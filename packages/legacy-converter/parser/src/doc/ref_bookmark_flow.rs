//! Bounded REF bookmark acquisition for the one Word-measured structural case:
//! a bookmarked page break followed by the referenced plain text.
//!
//! [MS-DOC] 2.8.10–2.8.12 (SttbfBkmk, PlcfBkf, PlcfBkl) locate the range;
//! ECMA-376 Part 1 §17.16.5.51 defines REF as bookmarked content. Word
//! controls with a leading break and a plain-text counterexample show that
//! the break precedes the cached REF result. Other bookmark structures keep
//! the cached result without inferred pagination.

use std::collections::{HashMap, HashSet};

use super::{fkp, u16_at, u32_at, unsupported};

const FC_NAMES: usize = 0x142;
const FC_STARTS: usize = 0x14a;
const FC_ENDS: usize = 0x152;
const MAX_BOOKMARKS: usize = 1024;
const MAX_TARGET_UNITS: usize = 4096;
const MAX_TOTAL_TARGET_BYTES: usize = 1024 * 1024;

pub(super) fn read(
    word: &[u8],
    table: &[u8],
    story: &str,
) -> Result<HashMap<String, String>, String> {
    let names = fkp::table_part(word, table, FC_NAMES)?;
    if names.is_empty() {
        return Ok(HashMap::new());
    }
    if names.len() < 6 || u16_at(names, 0)? != 0xffff || u16_at(names, 4)? != 0 {
        return Err(unsupported("invalid Word bookmark name table"));
    }
    let count = usize::from(u16_at(names, 2)?);
    if count > MAX_BOOKMARKS {
        return Err(unsupported("Word REF bookmark scan budget exceeded"));
    }
    let starts = fkp::table_part(word, table, FC_STARTS)?;
    let ends = fkp::table_part(word, table, FC_ENDS)?;
    if starts.len() != count.saturating_mul(8).saturating_add(4)
        || ends.len() != count.saturating_add(1).saturating_mul(4)
    {
        return Err(unsupported("invalid Word bookmark position arrays"));
    }
    let mut entries = Vec::with_capacity(count);
    let mut cursor = 6usize;
    let mut boundaries = Vec::with_capacity(count.saturating_mul(2));
    for index in 0..count {
        let units = usize::from(u16_at(names, cursor)?);
        if units == 0 || units >= 256 {
            return Err(unsupported("invalid Word bookmark name length"));
        }
        cursor = cursor
            .checked_add(2)
            .ok_or_else(|| unsupported("Word bookmark name overflow"))?;
        let end_name = cursor
            .checked_add(units.saturating_mul(2))
            .ok_or_else(|| unsupported("Word bookmark name overflow"))?;
        let name_bytes = names
            .get(cursor..end_name)
            .ok_or_else(|| unsupported("truncated Word bookmark name"))?;
        let name = String::from_utf16(
            &name_bytes
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
                .collect::<Vec<_>>(),
        )
        .map_err(|_| unsupported("invalid Word bookmark name"))?;
        cursor = end_name;
        let start = u32_at(starts, index * 4)? as usize;
        let end_index = usize::from(u16_at(starts, (count + 1) * 4 + index * 4)?);
        if end_index >= count {
            continue;
        }
        let end = u32_at(ends, end_index * 4)? as usize;
        if end <= start || end - start > MAX_TARGET_UNITS {
            continue;
        }
        let entry_index = entries.len();
        entries.push((name, start, end));
        boundaries.push((start, entry_index, true));
        boundaries.push((end, entry_index, false));
    }
    if cursor != names.len() {
        return Err(unsupported("Word bookmark name table has trailing bytes"));
    }
    boundaries.sort_unstable_by_key(|(cp, ..)| *cp);
    let mut offsets = vec![(None, None); entries.len()];
    let mut boundary_index = 0usize;
    let mut cp = 0usize;
    for (byte_offset, character) in story.char_indices() {
        while boundaries
            .get(boundary_index)
            .is_some_and(|(at, ..)| *at == cp)
        {
            let (_, entry, start) = boundaries[boundary_index];
            if start {
                offsets[entry].0 = Some(byte_offset);
            } else {
                offsets[entry].1 = Some(byte_offset);
            }
            boundary_index += 1;
        }
        cp += character.len_utf16();
    }
    while boundaries
        .get(boundary_index)
        .is_some_and(|(at, ..)| *at == cp)
    {
        let (_, entry, start) = boundaries[boundary_index];
        if start {
            offsets[entry].0 = Some(story.len());
        } else {
            offsets[entry].1 = Some(story.len());
        }
        boundary_index += 1;
    }
    let mut result = HashMap::new();
    let mut ambiguous = HashSet::new();
    let mut total = 0usize;
    for ((name, _, _), (start, end)) in entries.into_iter().zip(offsets) {
        let (Some(start), Some(end)) = (start, end) else {
            continue;
        };
        let Some(content) = story
            .get(start..end)
            .and_then(|text| text.strip_prefix("\u{c}\r"))
        else {
            continue;
        };
        if content.is_empty() || content.chars().any(char::is_control) {
            continue;
        }
        total = total.saturating_add(content.len());
        if total > MAX_TOTAL_TARGET_BYTES {
            break;
        }
        if !ambiguous.contains(&name) && result.insert(name.clone(), content.to_string()).is_some()
        {
            result.remove(&name);
            ambiguous.insert(name);
        }
    }
    Ok(result)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn bookmark_fixture(story: &str) -> (Vec<u8>, Vec<u8>) {
        let mut word = vec![0; 0x15a];
        let mut table = Vec::new();
        let mut names = vec![0xff, 0xff, 1, 0, 0, 0];
        names.extend_from_slice(&6u16.to_le_bytes());
        for unit in "target".encode_utf16() {
            names.extend_from_slice(&unit.to_le_bytes());
        }
        let end = story.encode_utf16().count() as u32;
        let starts = [0u32, end]
            .into_iter()
            .flat_map(u32::to_le_bytes)
            .chain([0u16, 0].into_iter().flat_map(u16::to_le_bytes))
            .collect::<Vec<_>>();
        let ends = [end, end]
            .into_iter()
            .flat_map(u32::to_le_bytes)
            .collect::<Vec<_>>();
        for (offset, part) in [(FC_NAMES, names), (FC_STARTS, starts), (FC_ENDS, ends)] {
            let fc = table.len() as u32;
            word[offset..offset + 4].copy_from_slice(&fc.to_le_bytes());
            word[offset + 4..offset + 8].copy_from_slice(&(part.len() as u32).to_le_bytes());
            table.extend(part);
        }
        (word, table)
    }

    #[test]
    fn only_one_leading_page_break_and_plain_text_is_replayed() {
        let story = "\u{c}\rTarget";
        let (word, table) = bookmark_fixture(story);
        assert_eq!(
            read(&word, &table, story).unwrap().get("target"),
            Some(&"Target".into())
        );
        for story in ["Target", "\u{c}\rTarget\rMore", "\u{c}\rTa\u{b}rget"] {
            let (word, table) = bookmark_fixture(story);
            assert!(read(&word, &table, story).unwrap().is_empty());
        }
    }
}
