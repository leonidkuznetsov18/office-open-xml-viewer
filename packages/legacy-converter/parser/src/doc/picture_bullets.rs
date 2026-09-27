//! The hidden main-story document containing Word picture bullets.
//! [MS-DOC] 2.6.1 sprmCPbiIBullet, 2.8.10-2.8.12 and 2.9.279.

use super::{fkp, u16_at, u32_at, unsupported};

const FC_NAMES: usize = 0x142; // FibRgFcLcb97.fcSttbfBkmk
const FC_STARTS: usize = 0x14a; // fcPlcfBkf
const FC_ENDS: usize = 0x152; // fcPlcfBkl

#[derive(Clone, Copy, Debug)]
pub(super) struct Document {
    start: usize,
    end: usize,
}

impl Document {
    pub(super) fn cp(self, relative: usize) -> Result<usize, String> {
        self.start
            .checked_add(relative)
            .filter(|cp| *cp < self.end)
            .ok_or_else(|| unsupported("Word picture bullet position outside hidden document"))
    }
}

/// Locate only `_PictureBullets`. Ordinary bookmarks are not projected, and
/// their text never becomes an image resource just because it is hidden.
/// Delaying this result until a picture marker is used keeps documents with
/// unrelated malformed bookmark tables on their established path.
pub(super) fn read(word: &[u8], table: &[u8]) -> Result<Option<Document>, String> {
    let names = fkp::table_part(word, table, FC_NAMES)?;
    if names.is_empty() {
        return Ok(None);
    }
    if names.len() < 6 || u16_at(names, 0)? != 0xffff || u16_at(names, 4)? != 0 {
        return Err(unsupported("invalid Word bookmark name table"));
    }
    let count = usize::from(u16_at(names, 2)?);
    if count > 0x3ffb {
        return Err(unsupported("Word bookmark count exceeds format limit"));
    }
    let mut cursor = 6usize;
    let mut selected = None;
    for index in 0..count {
        let units = usize::from(u16_at(names, cursor)?);
        if !(1..40).contains(&units) {
            return Err(unsupported("invalid Word bookmark name length"));
        }
        cursor = cursor
            .checked_add(2)
            .ok_or_else(|| unsupported("Word bookmark name overflow"))?;
        let bytes = units
            .checked_mul(2)
            .ok_or_else(|| unsupported("Word bookmark name overflow"))?;
        let raw = names
            .get(cursor..cursor + bytes)
            .ok_or_else(|| unsupported("truncated Word bookmark name"))?;
        if raw
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .eq("_PictureBullets".encode_utf16())
        {
            if selected.is_some() {
                return Err(unsupported("duplicate Word picture bullet bookmark"));
            }
            selected = Some(index);
        }
        cursor += bytes;
    }
    if cursor != names.len() {
        return Err(unsupported("Word bookmark name table has trailing bytes"));
    }
    let Some(index) = selected else {
        return Ok(None);
    };
    let starts = fkp::table_part(word, table, FC_STARTS)?;
    let ends = fkp::table_part(word, table, FC_ENDS)?;
    let start_bytes = count
        .checked_mul(8)
        .and_then(|bytes| bytes.checked_add(4))
        .ok_or_else(|| unsupported("Word bookmark array overflow"))?;
    let end_bytes = count
        .checked_add(1)
        .and_then(|count| count.checked_mul(4))
        .ok_or_else(|| unsupported("Word bookmark array overflow"))?;
    if starts.len() != start_bytes || ends.len() != end_bytes {
        return Err(unsupported("invalid Word bookmark position arrays"));
    }
    let start = u32_at(starts, index * 4)? as usize;
    let ibkl = usize::from(u16_at(starts, (count + 1) * 4 + index * 4)?);
    if ibkl >= count {
        return Err(unsupported(
            "Word picture bullet bookmark end index outside array",
        ));
    }
    let end = u32_at(ends, ibkl * 4)? as usize;
    let main_end = u32_at(word, super::CCP_TEXT_OFFSET)? as usize;
    if start >= end || end > main_end {
        return Err(unsupported(
            "Word picture bullet bookmark outside main document",
        ));
    }
    Ok(Some(Document { start, end }))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn hidden_bullet_bookmark_maps_relative_cp_through_end_index() {
        let mut word = vec![0; 0x15a];
        word[super::super::CCP_TEXT_OFFSET..super::super::CCP_TEXT_OFFSET + 4]
            .copy_from_slice(&100u32.to_le_bytes());
        let mut table = Vec::new();
        let names = {
            let mut bytes = vec![0xff, 0xff, 2, 0, 0, 0];
            for name in ["ordinary", "_PictureBullets"] {
                bytes.extend_from_slice(&(name.len() as u16).to_le_bytes());
                for unit in name.encode_utf16() {
                    bytes.extend_from_slice(&unit.to_le_bytes());
                }
            }
            bytes
        };
        let starts = [10u32, 90, 100]
            .into_iter()
            .flat_map(u32::to_le_bytes)
            .chain([1u16, 0, 0, 0].into_iter().flat_map(u16::to_le_bytes))
            .collect::<Vec<_>>();
        let ends = [92u32, 12, 100]
            .into_iter()
            .flat_map(u32::to_le_bytes)
            .collect::<Vec<_>>();
        for (offset, part) in [(FC_NAMES, names), (FC_STARTS, starts), (FC_ENDS, ends)] {
            let fc = table.len() as u32;
            word[offset..offset + 4].copy_from_slice(&fc.to_le_bytes());
            word[offset + 4..offset + 8].copy_from_slice(&(part.len() as u32).to_le_bytes());
            table.extend(part);
        }
        let document = read(&word, &table).unwrap().unwrap();
        assert_eq!(document.cp(0).unwrap(), 90);
        assert_eq!(document.cp(1).unwrap(), 91);
        assert!(document.cp(2).is_err());
    }
}
