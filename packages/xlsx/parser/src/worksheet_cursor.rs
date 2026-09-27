//! Production lifecycle for resumable worksheet-row projection.
//!
//! The cursor owns the worksheet entry stream and projector between pulls. The
//! surrounding [`XlsxZip`](crate::XlsxZip) continues to own the single logical
//! package operation so workbook dependencies, every row pull, and the sheet's
//! ancillary parts are charged to the same operation.

use std::borrow::Cow;
use std::collections::BTreeMap;
use std::io::{Cursor, Read};
use std::rc::Rc;

#[cfg(test)]
use ooxml_common::bounded_xml::BoundedXmlReader;
use ooxml_common::bounded_xml::MCE_NS;
#[cfg(test)]
use ooxml_common::ns::is_x_ns;
#[cfg(test)]
use quick_xml::events::{BytesStart, Event};

use crate::worksheet_projector::{
    ProjectedWorksheetRow, WorksheetProjectorItem, WorksheetRowProjector,
};
use crate::{
    parse_cell_ref_checked, resolve_implicit_ordinal, Row, SharedString, SpreadsheetOrdinal,
    XlsxZip,
};

/// Default semantic credit for one production pull. Rows are indivisible: the
/// cursor never splits a row to satisfy this credit.
pub(super) const WORKSHEET_CURSOR_PULL_ROWS: usize = 128;
pub(super) const WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES: usize = 1024 * 1024;

type ProductionProjector =
    WorksheetRowProjector<std::io::BufReader<Box<dyn Read>>, Rc<[SharedString]>, Rc<[String]>>;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum WorksheetCursorState {
    Open,
    Finished,
    Failed,
    Canceled,
    Closed,
}

#[derive(Debug)]
pub(super) struct WorksheetCursorTail {
    pub(super) shell_xml: String,
    pub(super) row_heights: BTreeMap<u32, f64>,
}

/// Facts that the viewer needs before accepting an exact provisional frame.
/// The scan observes the MCE-processed infoset and keeps no cell values.
pub(super) struct WorksheetCursorPreview {
    pub(super) tail: Option<WorksheetCursorTail>,
    pub(super) raw: Rc<[u8]>,
    pub(super) max_row: u32,
    pub(super) max_col: u32,
    pub(super) has_row_outline: bool,
    pub(super) ordered_rows: bool,
}

#[derive(Debug)]
pub(super) enum WorksheetCursorPull {
    Rows {
        rows: Vec<Row>,
        /// Sum of internal standalone row-projection bytes. This is not a
        /// serialized wire size and therefore is not protocol byte credit.
        projected_bytes: usize,
    },
    Finished(WorksheetCursorTail),
}

/// An owned worksheet-entry cursor whose projector survives across pulls.
///
/// Row batches are provisional until [`WorksheetCursorPull::Finished`]. A
/// caller assembling retained state must not commit earlier batches: the ZIP
/// CRC and well-formed worksheet tail are validated only when the entry reaches
/// EOF. `cancel` and `close` are deliberately idempotent and immediately drop
/// the entry stream/decoder lease.
pub(super) struct WorksheetCursor {
    projector: Option<ProductionProjector>,
    pending_row: Option<ProjectedWorksheetRow>,
    pending_tail: Option<WorksheetCursorTail>,
    state: WorksheetCursorState,
}

impl WorksheetCursor {
    fn open_under_active_operation(
        archive: &mut XlsxZip,
        part: &str,
        shared_strings: Rc<[SharedString]>,
        theme_colors: Rc<[String]>,
    ) -> Result<Self, String> {
        let entry = archive.active_operation()?.open_entry(part)?;
        let reporter = entry.limit_reporter()?;
        let projector = WorksheetRowProjector::from_owned_reader(
            Box::new(entry),
            part.to_string(),
            reporter,
            shared_strings,
            theme_colors,
        );
        Ok(Self {
            projector: Some(projector),
            pending_row: None,
            pending_tail: None,
            state: WorksheetCursorState::Open,
        })
    }

    fn open_from_buffer_under_active_operation(
        archive: &mut XlsxZip,
        part: &str,
        raw: Rc<[u8]>,
        shared_strings: Rc<[SharedString]>,
        theme_colors: Rc<[String]>,
    ) -> Result<Self, String> {
        let reporter = archive.active_operation()?.limit_reporter()?;
        let projector = WorksheetRowProjector::from_owned_reader(
            Box::new(Cursor::new(raw)),
            part.to_string(),
            reporter,
            shared_strings,
            theme_colors,
        );
        Ok(Self {
            projector: Some(projector),
            pending_row: None,
            pending_tail: None,
            state: WorksheetCursorState::Open,
        })
    }

    pub(super) fn pull(
        &mut self,
        max_rows: usize,
        target_projected_bytes: usize,
    ) -> Result<WorksheetCursorPull, String> {
        // `max_rows` is a hard semantic-unit limit and is clamped below. The
        // projected-byte argument is intentionally a soft internal batching
        // target: one indivisible row may cross it after passing the separate
        // 8 MiB hard row-projection cap. A future wire adapter must measure its
        // serialized payload independently and obey protocol `byteCredit`.
        if max_rows == 0 || target_projected_bytes == 0 {
            return Err("worksheet cursor pull limits must be greater than zero".to_string());
        }
        if self.state != WorksheetCursorState::Open {
            return Err(self.inactive_error());
        }
        if let Some(tail) = self.pending_tail.take() {
            self.state = WorksheetCursorState::Finished;
            self.projector.take();
            return Ok(WorksheetCursorPull::Finished(tail));
        }

        let row_limit = max_rows.min(WORKSHEET_CURSOR_PULL_ROWS);
        let mut rows = Vec::with_capacity(row_limit);
        let mut projected_bytes = 0usize;
        loop {
            let item = match self.pending_row.take() {
                Some(row) => Ok(WorksheetProjectorItem::Row(row)),
                None => self
                    .projector
                    .as_mut()
                    .expect("open worksheet cursor owns its projector")
                    .next_item(),
            };
            match item {
                Ok(WorksheetProjectorItem::Row(row)) => {
                    let next_bytes = projected_bytes.saturating_add(row.projected_bytes);
                    if !rows.is_empty() && next_bytes > target_projected_bytes {
                        self.pending_row = Some(row);
                        return Ok(WorksheetCursorPull::Rows {
                            rows,
                            projected_bytes,
                        });
                    }
                    projected_bytes = next_bytes;
                    rows.push(row.row);
                    if rows.len() == row_limit {
                        return Ok(WorksheetCursorPull::Rows {
                            rows,
                            projected_bytes,
                        });
                    }
                }
                Ok(WorksheetProjectorItem::Finished(tail)) => {
                    let tail = WorksheetCursorTail {
                        shell_xml: tail.shell_xml,
                        row_heights: tail.row_heights,
                    };
                    if rows.is_empty() {
                        self.state = WorksheetCursorState::Finished;
                        self.projector.take();
                        return Ok(WorksheetCursorPull::Finished(tail));
                    }
                    self.pending_tail = Some(tail);
                    return Ok(WorksheetCursorPull::Rows {
                        rows,
                        projected_bytes,
                    });
                }
                Err(error) => {
                    self.state = WorksheetCursorState::Failed;
                    self.projector.take();
                    self.pending_tail = None;
                    return Err(error.to_string());
                }
            }
        }
    }

    pub(super) fn cancel(&mut self) {
        if matches!(
            self.state,
            WorksheetCursorState::Canceled | WorksheetCursorState::Closed
        ) {
            return;
        }
        self.projector.take();
        self.pending_row = None;
        self.pending_tail = None;
        self.state = WorksheetCursorState::Canceled;
    }

    pub(super) fn close(&mut self) {
        if self.state == WorksheetCursorState::Closed {
            return;
        }
        self.projector.take();
        self.pending_row = None;
        self.pending_tail = None;
        self.state = WorksheetCursorState::Closed;
    }

    fn inactive_error(&self) -> String {
        let state = match self.state {
            WorksheetCursorState::Open => "open",
            WorksheetCursorState::Finished => "finished",
            WorksheetCursorState::Failed => "failed",
            WorksheetCursorState::Canceled => "canceled",
            WorksheetCursorState::Closed => "closed",
        };
        format!("worksheet cursor is {state}")
    }
}

impl XlsxZip {
    pub(super) fn scan_worksheet_preview(
        &mut self,
        part: &str,
        shared_strings: Rc<[SharedString]>,
        theme_colors: Rc<[String]>,
    ) -> Result<WorksheetCursorPreview, String> {
        let mut entry = self.active_operation()?.open_entry(part)?;
        let mut bytes = Vec::new();
        entry
            .read_to_end(&mut bytes)
            .map_err(|error| error.to_string())?;
        let raw: Rc<[u8]> = bytes.into();
        let scanned = lexical_scan_worksheet_preview(Rc::clone(&raw));
        let mut preview = match scanned {
            Ok(preview) => preview,
            Err(_) => {
                return Ok(WorksheetCursorPreview {
                    tail: None,
                    raw,
                    max_row: 0,
                    max_col: 0,
                    has_row_outline: false,
                    ordered_rows: false,
                })
            }
        };
        if let Some(tail) = preview.tail.as_mut().filter(|_| {
            raw.windows(MCE_NS.len())
                .any(|window| window == MCE_NS.as_bytes())
        }) {
            // The shell may carry MCE attributes. Apply the same projection as
            // the terminal cursor, but only to the small, row-free shell.
            let shell = tail.shell_xml.as_bytes().to_vec();
            let reporter = self.active_operation()?.limit_reporter()?;
            let mut projector = WorksheetRowProjector::from_owned_reader(
                Box::new(Cursor::new(shell)),
                part.to_string(),
                reporter,
                shared_strings,
                theme_colors,
            );
            match projector.next_item().map_err(|error| error.to_string())? {
                WorksheetProjectorItem::Finished(projected) => {
                    tail.shell_xml = projected.shell_xml;
                }
                WorksheetProjectorItem::Row(_) => {
                    return Err("row-free worksheet shell contained a row".to_string());
                }
            }
        }
        Ok(preview)
    }

    /// Open a persistent worksheet cursor only inside an explicitly-started
    /// package operation. This makes operation ownership visible at the factory
    /// boundary and prevents the lazy compatibility operation from escaping
    /// across production pulls.
    pub(super) fn open_worksheet_cursor(
        &mut self,
        part: &str,
        shared_strings: Rc<[SharedString]>,
        theme_colors: Rc<[String]>,
    ) -> Result<WorksheetCursor, String> {
        WorksheetCursor::open_under_active_operation(self, part, shared_strings, theme_colors)
    }

    pub(super) fn open_buffered_worksheet_cursor(
        &mut self,
        part: &str,
        raw: Rc<[u8]>,
        shared_strings: Rc<[SharedString]>,
        theme_colors: Rc<[String]>,
    ) -> Result<WorksheetCursor, String> {
        WorksheetCursor::open_from_buffer_under_active_operation(
            self,
            part,
            raw,
            shared_strings,
            theme_colors,
        )
    }
}

#[cfg(test)]
fn numeric_attribute(start: &BytesStart<'_>, name: &[u8]) -> Result<Option<String>, String> {
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| error.to_string())?;
        if attribute.key.as_ref() == name {
            return std::str::from_utf8(attribute.value.as_ref())
                .map(str::to_string)
                .map(Some)
                .map_err(|error| error.to_string());
        }
    }
    Ok(None)
}

/// Read the tail and row coordinates without projecting cell bodies. The
/// ordinary cursor later validates and projects the same retained bytes. MCE
/// worksheets stay on that cursor's complete, MCE-aware path.
#[cfg(test)]
fn fast_scan_worksheet_preview(raw: Rc<[u8]>) -> Result<WorksheetCursorPreview, String> {
    let mut reader = BoundedXmlReader::new(
        std::io::BufReader::new(Cursor::new(Rc::clone(&raw))),
        1024 * 1024,
        "worksheet preview",
    );
    let mut sheet_start = None;
    let mut sheet_end = None;
    let mut inside_sheet = false;
    let mut previous_row = 0;
    let mut previous_col = 0;
    let mut max_row = 0;
    let mut max_col = 0;
    let mut has_row_outline = false;
    let mut ordered_rows = true;
    let mut row_heights = BTreeMap::new();
    loop {
        let read = reader.read_event().map_err(|error| format!("{error:?}"))?;
        let x = is_x_ns(read.namespace.as_deref());
        let empty = matches!(read.event, Event::Empty(_));
        match read.event {
            Event::Start(start) | Event::Empty(start)
                if x && start.local_name().as_ref() == b"sheetData" =>
            {
                if sheet_start.is_some() {
                    return Err("worksheet has repeated sheetData".to_string());
                }
                sheet_start = Some(read.span.end as usize);
                if empty {
                    sheet_end = sheet_start;
                } else {
                    inside_sheet = true;
                }
            }
            Event::End(end) if x && end.local_name().as_ref() == b"sheetData" => {
                sheet_end = Some(read.span.start as usize);
                inside_sheet = false;
            }
            Event::Start(start) | Event::Empty(start)
                if inside_sheet && x && start.local_name().as_ref() == b"row" =>
            {
                let explicit = numeric_attribute(&start, b"r")?
                    .map(|value| {
                        value
                            .parse::<u32>()
                            .map_err(|_| format!("invalid row ordinal: {value}"))
                    })
                    .transpose()?;
                let previous = previous_row;
                let row =
                    resolve_implicit_ordinal(explicit, &mut previous_row, SpreadsheetOrdinal::Row)?;
                ordered_rows &= row > previous;
                max_row = max_row.max(row);
                previous_col = 0;
                let hidden = matches!(
                    numeric_attribute(&start, b"hidden")?.as_deref(),
                    Some("1" | "true")
                );
                let height = if hidden {
                    Some(0.0)
                } else {
                    numeric_attribute(&start, b"ht")?
                        .and_then(|value| value.parse::<f64>().ok())
                        .filter(|value| value.is_finite() && *value >= 0.0)
                };
                if let Some(height) = height {
                    row_heights.insert(row, height);
                }
                let outline = numeric_attribute(&start, b"outlineLevel")?
                    .and_then(|value| value.parse::<u8>().ok())
                    .unwrap_or(0);
                let collapsed = matches!(
                    numeric_attribute(&start, b"collapsed")?.as_deref(),
                    Some("1" | "true")
                );
                has_row_outline |= outline != 0 || collapsed;
            }
            Event::Start(start) | Event::Empty(start)
                if inside_sheet && x && start.local_name().as_ref() == b"c" =>
            {
                let explicit = numeric_attribute(&start, b"r")?
                    .map(|reference| parse_cell_ref_checked(&reference).map(|(col, _)| col))
                    .transpose()?;
                let col = resolve_implicit_ordinal(
                    explicit,
                    &mut previous_col,
                    SpreadsheetOrdinal::Column,
                )?;
                max_col = max_col.max(col);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let (Some(start), Some(end)) = (sheet_start, sheet_end) else {
        return Err("worksheet has no complete sheetData".to_string());
    };
    if end < start || end > raw.len() {
        return Err("worksheet sheetData boundary is invalid".to_string());
    }
    if raw[start..end]
        .windows(b"AlternateContent".len())
        .any(|window| window == b"AlternateContent")
        || raw
            .windows(b"ProcessContent".len())
            .any(|window| window == b"ProcessContent")
    {
        return Err("worksheet row MCE requires the complete projector".to_string());
    }
    let mut shell = Vec::with_capacity(raw.len() - (end - start));
    shell.extend_from_slice(&raw[..start]);
    shell.extend_from_slice(&raw[end..]);
    let shell_xml = String::from_utf8(shell).map_err(|error| error.to_string())?;
    Ok(WorksheetCursorPreview {
        tail: Some(WorksheetCursorTail {
            shell_xml,
            row_heights,
        }),
        raw,
        max_row,
        max_col,
        has_row_outline,
        ordered_rows,
    })
}

fn find_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

fn tag_end(bytes: &[u8], start: usize) -> Result<usize, String> {
    let mut quote = 0;
    for (offset, byte) in bytes[start..].iter().enumerate() {
        if offset > 1024 * 1024 {
            return Err("worksheet preview tag exceeds event limit".to_string());
        }
        if quote == 0 {
            match byte {
                b'\'' | b'"' => quote = *byte,
                b'>' => return Ok(start + offset + 1),
                _ => {}
            }
        } else if *byte == quote {
            quote = 0;
        }
    }
    Err("worksheet preview tag is unclosed".to_string())
}

fn tag_attribute<'a>(tag: &'a [u8], name: &[u8]) -> Result<Option<Cow<'a, str>>, String> {
    let mut i = 1;
    while i < tag.len() && !tag[i].is_ascii_whitespace() && tag[i] != b'>' && tag[i] != b'/' {
        i += 1;
    }
    while i < tag.len() {
        while i < tag.len() && tag[i].is_ascii_whitespace() {
            i += 1;
        }
        if i == tag.len() || tag[i] == b'/' || tag[i] == b'>' {
            return Ok(None);
        }
        let key_start = i;
        while i < tag.len() && !tag[i].is_ascii_whitespace() && tag[i] != b'=' {
            i += 1;
        }
        let key = &tag[key_start..i];
        while i < tag.len() && tag[i].is_ascii_whitespace() {
            i += 1;
        }
        if i == tag.len() || tag[i] != b'=' {
            return Err("worksheet preview attribute is malformed".to_string());
        }
        i += 1;
        while i < tag.len() && tag[i].is_ascii_whitespace() {
            i += 1;
        }
        if i == tag.len() || (tag[i] != b'\'' && tag[i] != b'"') {
            return Err("worksheet preview attribute is unquoted".to_string());
        }
        let quote = tag[i];
        i += 1;
        let value_start = i;
        while i < tag.len() && tag[i] != quote {
            i += 1;
        }
        if i == tag.len() {
            return Err("worksheet preview attribute is unclosed".to_string());
        }
        let value = &tag[value_start..i];
        i += 1;
        if key == name {
            let encoded = std::str::from_utf8(value).map_err(|error| error.to_string())?;
            if encoded.contains('<') {
                return Err("worksheet preview attribute contains unescaped markup".to_string());
            }
            if !encoded.contains('&') {
                return Ok(Some(Cow::Borrowed(encoded)));
            }
            // The terminal row parser reads roxmltree's decoded attribute
            // values. Use that same XML decoder for the uncommon entity path;
            // a malformed or unsupported entity disables the preview.
            let delimiter = quote as char;
            let xml = format!("<v a={delimiter}{encoded}{delimiter}/>");
            let document = roxmltree::Document::parse(&xml).map_err(|error| error.to_string())?;
            let decoded = document
                .root_element()
                .attribute("a")
                .ok_or_else(|| "worksheet preview attribute disappeared".to_string())?;
            return Ok(Some(Cow::Owned(decoded.to_string())));
        }
    }
    Ok(None)
}

fn lexical_scan_worksheet_preview(raw: Rc<[u8]>) -> Result<WorksheetCursorPreview, String> {
    // The fast path is restricted to ordinary unprefixed SpreadsheetML rows.
    // Comments, CDATA and MCE row replacement use the complete projector.
    let Some(open) = find_bytes(&raw, b"<sheetData") else {
        return Err("worksheet has no unprefixed sheetData".to_string());
    };
    if raw
        .get(open + b"<sheetData".len())
        .is_some_and(|byte| !byte.is_ascii_whitespace() && *byte != b'/' && *byte != b'>')
    {
        return Err("worksheet sheetData tag is ambiguous".to_string());
    }
    let start = tag_end(&raw, open)?;
    let self_closing = raw[open..start].ends_with(b"/>");
    let end = if self_closing {
        start
    } else {
        let close = find_bytes(&raw[start..], b"</sheetData>")
            .ok_or_else(|| "worksheet has no closing sheetData".to_string())?;
        start + close
    };
    let body = &raw[start..end];
    if find_bytes(&raw, b"<!--").is_some()
        || find_bytes(&raw, b"<![CDATA[").is_some()
        || find_bytes(body, b"<?").is_some()
        || find_bytes(body, b"AlternateContent").is_some()
        || find_bytes(&raw, b"ProcessContent").is_some()
    {
        return Err("worksheet requires the complete XML projector".to_string());
    }
    let mut previous_row = 0;
    let mut previous_col = 0;
    let mut max_row = 0;
    let mut max_col = 0;
    let mut has_row_outline = false;
    let mut ordered_rows = true;
    let mut row_heights = BTreeMap::new();
    let mut offset = 0;
    while let Some(found) = body[offset..].iter().position(|byte| *byte == b'<') {
        let begin = offset + found;
        let after = begin + 1;
        if after >= body.len() {
            return Err("worksheet preview ends in a tag".to_string());
        }
        let close = tag_end(body, begin)?;
        offset = close;
        if body[after] == b'/' {
            continue;
        }
        if body[after] == b'!' || body[after] == b'?' {
            return Err("worksheet preview has special row markup".to_string());
        }
        let mut name_end = after;
        while name_end < close
            && !body[name_end].is_ascii_whitespace()
            && body[name_end] != b'/'
            && body[name_end] != b'>'
        {
            name_end += 1;
        }
        let name = &body[after..name_end];
        if name.contains(&b':') {
            return Err("worksheet preview has prefixed row markup".to_string());
        }
        let tag = &body[begin..close];
        if name == b"row" {
            let explicit = tag_attribute(tag, b"r")?
                .map(|value| value.parse::<u32>().map_err(|error| error.to_string()))
                .transpose()?;
            let prior = previous_row;
            let row =
                resolve_implicit_ordinal(explicit, &mut previous_row, SpreadsheetOrdinal::Row)?;
            ordered_rows &= row > prior;
            max_row = max_row.max(row);
            previous_col = 0;
            let hidden = matches!(tag_attribute(tag, b"hidden")?, Some(value) if value == "1" || value == "true");
            let height = if hidden {
                Some(0.0)
            } else {
                tag_attribute(tag, b"ht")?
                    .and_then(|value| value.parse::<f64>().ok())
                    .filter(|value| value.is_finite() && *value >= 0.0)
            };
            if let Some(height) = height {
                row_heights.insert(row, height);
            }
            let outline = tag_attribute(tag, b"outlineLevel")?
                .and_then(|value| value.parse::<u8>().ok())
                .unwrap_or(0);
            let collapsed = matches!(tag_attribute(tag, b"collapsed")?, Some(value) if value == "1" || value == "true");
            has_row_outline |= outline != 0 || collapsed;
        } else if name == b"c" {
            let explicit = tag_attribute(tag, b"r")?
                .map(|reference| parse_cell_ref_checked(&reference).map(|(col, _)| col))
                .transpose()?;
            let col =
                resolve_implicit_ordinal(explicit, &mut previous_col, SpreadsheetOrdinal::Column)?;
            max_col = max_col.max(col);
        }
    }
    let mut shell = Vec::with_capacity(raw.len() - (end - start));
    shell.extend_from_slice(&raw[..start]);
    shell.extend_from_slice(&raw[end..]);
    if shell.len() > 16 * 1024 * 1024 {
        return Err("worksheet preview shell exceeds retained limit".to_string());
    }
    let shell_xml = String::from_utf8(shell).map_err(|error| error.to_string())?;
    Ok(WorksheetCursorPreview {
        tail: Some(WorksheetCursorTail {
            shell_xml,
            row_heights,
        }),
        raw,
        max_row,
        max_col,
        has_row_outline,
        ordered_rows,
    })
}

impl Drop for WorksheetCursor {
    fn drop(&mut self) {
        self.close();
    }
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use zip::write::SimpleFileOptions;

    use super::*;
    use crate::open_zip;

    const PART: &str = "xl/worksheets/sheet1.xml";

    fn worksheet(row_count: usize, tail: &str) -> String {
        let mut xml = String::from(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
        );
        for row in 1..=row_count {
            xml.push_str(&format!(
                r#"<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>row {row}</t></is></c></row>"#
            ));
        }
        xml.push_str(tail);
        xml
    }

    fn package(xml: &str) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(
                    PART,
                    SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored),
                )
                .unwrap();
            writer.write_all(xml.as_bytes()).unwrap();
            writer.finish().unwrap();
        }
        bytes
    }

    fn corrupt_crc_consistently(bytes: &mut [u8]) {
        let wrong_crc = u32::from_le_bytes(bytes[14..18].try_into().unwrap()) ^ 0xffff_ffff;
        bytes[14..18].copy_from_slice(&wrong_crc.to_le_bytes());
        let central = bytes
            .windows(4)
            .position(|window| window == 0x0201_4b50u32.to_le_bytes())
            .expect("central directory header");
        bytes[central + 16..central + 20].copy_from_slice(&wrong_crc.to_le_bytes());
    }

    #[test]
    fn lexical_tail_scan_matches_xml_scan_for_implicit_and_explicit_coordinates() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2" ht="21"><c r="B2"><v>1</v></c><c><v>2</v></c></row><row hidden="1"><c r="D3"/></row><row r="8" outlineLevel="2"><c r="AA8"><v>3</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="B2:C2"/></mergeCells></worksheet>"#;
        let raw: Rc<[u8]> = xml.as_bytes().into();
        let lexical = lexical_scan_worksheet_preview(Rc::clone(&raw)).unwrap();
        let parsed = fast_scan_worksheet_preview(raw).unwrap();
        assert_eq!(lexical.max_row, parsed.max_row);
        assert_eq!(lexical.max_col, parsed.max_col);
        assert_eq!(lexical.has_row_outline, parsed.has_row_outline);
        assert_eq!(lexical.ordered_rows, parsed.ordered_rows);
        assert_eq!(
            lexical.tail.unwrap().row_heights,
            parsed.tail.unwrap().row_heights
        );
    }

    #[test]
    fn lexical_scan_decodes_row_attributes_like_the_terminal_xml_parser() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="&#50;" ht="&#51;0"><c r="B&#50;"><v>1</v></c></row><row r="3" ht="3&amp;0"><c r="C3"><v>2</v></c></row></sheetData></worksheet>"#;
        let preview = lexical_scan_worksheet_preview(xml.as_bytes().into()).unwrap();
        assert_eq!(preview.max_row, 3);
        assert_eq!(preview.max_col, 3);
        assert_eq!(
            preview.tail.unwrap().row_heights,
            BTreeMap::from([(2, 30.0)])
        );

        // The main row projector uses roxmltree attribute values. Compare the
        // same decoded values, including a predefined named entity whose
        // result is not a valid row height.
        let document = roxmltree::Document::parse(xml).unwrap();
        let rows: Vec<_> = document
            .descendants()
            .filter(|node| node.has_tag_name("row"))
            .collect();
        assert_eq!(rows[0].attribute("ht"), Some("30"));
        assert_eq!(rows[1].attribute("ht"), Some("3&0"));

        let unsupported = xml.replace("3&amp;0", "3&unknown;0");
        assert!(lexical_scan_worksheet_preview(unsupported.as_bytes().into()).is_err());
    }

    #[test]
    fn production_cursor_keeps_one_operation_across_multiple_atomic_pulls() {
        let xml = worksheet(600, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).expect("package opens");
        archive
            .begin_operation("worksheet-cursor")
            .expect("operation starts");
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .expect("cursor opens");

        let mut batches = Vec::new();
        let mut inflated_snapshots = Vec::new();
        let (batches, tail) = loop {
            match cursor
                .pull(64, WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES)
                .expect("pull succeeds")
            {
                WorksheetCursorPull::Rows { rows, .. } => {
                    assert!(!rows.is_empty());
                    assert!(rows.len() <= 64);
                    batches.push(rows);
                    inflated_snapshots.push(
                        archive
                            .operation
                            .active()
                            .unwrap()
                            .usage()
                            .unwrap()
                            .operation_inflated_bytes,
                    );
                }
                WorksheetCursorPull::Finished(tail) => {
                    break (batches, tail);
                }
            }
        };

        assert_eq!(batches.iter().map(Vec::len).sum::<usize>(), 600);
        assert_eq!(
            batches
                .iter()
                .flatten()
                .map(|row| row.index)
                .collect::<Vec<_>>(),
            (1..=600).collect::<Vec<_>>()
        );
        assert!(
            inflated_snapshots.windows(2).any(|pair| pair[0] < pair[1]),
            "the same operation continues inflating after earlier row pulls"
        );
        assert!(tail.shell_xml.contains("<sheetData></sheetData>"));
        assert!(archive.operation.active().unwrap().usage().is_some());
        archive.finish_operation().expect("same operation finishes");
    }

    #[test]
    fn late_malformed_tail_never_produces_a_committable_transaction() {
        let xml = worksheet(600, "</sheetData><broken>");
        let mut archive = open_zip(package(&xml)).expect("package opens");
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        let mut provisional = Vec::new();
        let mut finished = false;

        loop {
            match cursor.pull(7, WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES) {
                Ok(WorksheetCursorPull::Rows { rows, .. }) => provisional.extend(rows),
                Ok(WorksheetCursorPull::Finished(_)) => {
                    finished = true;
                    break;
                }
                Err(error) => {
                    assert!(error.contains("EOF") || error.contains("closed"), "{error}");
                    break;
                }
            }
        }

        assert!(
            !provisional.is_empty(),
            "earlier pulls are intentionally provisional"
        );
        assert!(!finished, "malformed tail must prevent commit");
        archive.cancel_operation();
    }

    #[test]
    fn crc_failure_after_row_pulls_never_produces_finished() {
        let xml = worksheet(600, "</sheetData></worksheet>");
        let mut bytes = package(&xml);
        corrupt_crc_consistently(&mut bytes);
        let mut archive = open_zip(bytes).expect("matching forged metadata passes preflight");
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        let mut provisional_rows = 0;

        let error = loop {
            match cursor.pull(7, WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES) {
                Ok(WorksheetCursorPull::Rows { rows, .. }) => provisional_rows += rows.len(),
                Ok(WorksheetCursorPull::Finished(_)) => panic!("CRC failure must prevent commit"),
                Err(error) => break error,
            }
        };

        assert!(provisional_rows > 0);
        assert!(error.contains("CRC"), "{error}");
        archive.cancel_operation();
    }

    #[test]
    fn close_and_cancel_are_idempotent_and_release_the_entry() {
        let xml = worksheet(2, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        archive.begin_operation("parse-sheet").unwrap();

        let mut closed = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        closed.close();
        closed.close();
        assert_eq!(closed.pull(1, 1).unwrap_err(), "worksheet cursor is closed");

        let mut canceled = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        canceled.cancel();
        canceled.cancel();
        assert_eq!(
            canceled.pull(1, 1).unwrap_err(),
            "worksheet cursor is canceled"
        );
        archive
            .finish_operation()
            .expect("released readers allow finish");
    }

    #[test]
    fn pull_clamps_rows_and_stages_the_first_soft_projection_overrun() {
        let xml = worksheet(600, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();

        let first = cursor.pull(usize::MAX, usize::MAX).unwrap();
        let WorksheetCursorPull::Rows { rows, .. } = first else {
            panic!("large worksheet must yield rows");
        };
        assert_eq!(rows.len(), WORKSHEET_CURSOR_PULL_ROWS);

        let mut observed = rows.into_iter().map(|row| row.index).collect::<Vec<_>>();
        let mut saw_multi_row_projection_limited_batch = false;
        while let WorksheetCursorPull::Rows {
            rows,
            projected_bytes,
        } = cursor.pull(usize::MAX, 500).unwrap()
        {
            assert!(!rows.is_empty());
            if rows.len() > 1 {
                saw_multi_row_projection_limited_batch = true;
                assert!(projected_bytes <= 500);
            }
            observed.extend(rows.into_iter().map(|row| row.index));
        }
        assert!(saw_multi_row_projection_limited_batch);
        assert_eq!(observed, (1..=600).collect::<Vec<_>>());
        archive.finish_operation().unwrap();
    }

    #[test]
    fn indivisible_row_may_cross_soft_projection_target_but_not_hard_row_cap() {
        let xml = worksheet(2, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();

        for expected in 1..=2 {
            let WorksheetCursorPull::Rows {
                rows,
                projected_bytes,
            } = cursor.pull(128, 1).unwrap()
            else {
                panic!("row is returned atomically");
            };
            assert_eq!(rows.len(), 1);
            assert_eq!(rows[0].index, expected);
            assert!(projected_bytes > 1);
            assert!(projected_bytes <= crate::worksheet_projector::STREAMED_ROW_PROJECTION_BYTES);
        }
        assert!(matches!(
            cursor.pull(128, 1).unwrap(),
            WorksheetCursorPull::Finished(_)
        ));
        archive.finish_operation().unwrap();
    }

    #[test]
    fn cursor_factory_requires_an_explicit_active_operation() {
        let xml = worksheet(1, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        let error = match archive.open_worksheet_cursor(PART, Rc::from([]), Rc::from([])) {
            Ok(_) => panic!("cursor factory must not create a compatibility operation"),
            Err(error) => error,
        };
        assert_eq!(error, "xlsx package operation is not active");
        assert!(!archive.operation.is_active());
    }
}
