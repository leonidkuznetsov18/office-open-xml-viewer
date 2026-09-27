use std::any::Any;
use std::collections::{HashMap, VecDeque};
use std::fs;
use std::io::Read;
use std::path::PathBuf;
use std::sync::{Arc, Mutex, OnceLock};
use std::time::SystemTime;

use docx_model::Document;
use ooxml_common::resource::HARD_MAX_TOTAL_INFLATED_BYTES;
use pptx_model::Presentation;
use xlsx_model::{Workbook, Worksheet};

const CACHE_ENTRIES: usize = 8;

#[derive(Clone, PartialEq, Eq)]
struct Key {
    path: PathBuf,
    modified: SystemTime,
    len: u64,
}

type CachedModel = Arc<dyn Any + Send + Sync>;
struct Entry {
    key: Key,
    model: Option<CachedModel>,
    markdown: HashMap<MarkdownKind, Arc<String>>,
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub enum MarkdownKind {
    Docx,
    Xlsx,
    Pptx,
}

static CACHE: OnceLock<Mutex<VecDeque<Entry>>> = OnceLock::new();

fn cache() -> &'static Mutex<VecDeque<Entry>> {
    CACHE.get_or_init(|| Mutex::new(VecDeque::new()))
}

fn identity(path: &str) -> Result<Key, String> {
    let canonical = fs::canonicalize(path).map_err(|e| format!("Cannot read '{}': {}", path, e))?;
    let meta = fs::metadata(&canonical).map_err(|e| format!("Cannot read '{}': {}", path, e))?;
    if !meta.is_file() {
        return Err(format!("Cannot read '{}': not a regular file", path));
    }
    // The shared hard total-inflated-byte ceiling is also an upper bound on
    // the compressed package accepted for eager MCP reads. This is MCP input
    // governance; it does not alter parser or WASM limits.
    if meta.len() > HARD_MAX_TOTAL_INFLATED_BYTES {
        return Err(format!(
            "Cannot read '{}': package exceeds {} bytes",
            path, HARD_MAX_TOTAL_INFLATED_BYTES
        ));
    }
    Ok(Key {
        path: canonical,
        modified: meta
            .modified()
            .map_err(|e| format!("Cannot read '{}': {}", path, e))?,
        len: meta.len(),
    })
}

fn read_checked(path: &str, key: &Key) -> Result<Vec<u8>, String> {
    let file = fs::File::open(&key.path).map_err(|e| format!("Cannot read '{}': {}", path, e))?;
    let mut data = Vec::new();
    file.take(HARD_MAX_TOTAL_INFLATED_BYTES + 1)
        .read_to_end(&mut data)
        .map_err(|e| format!("Cannot read '{}': {}", path, e))?;
    if data.len() as u64 > HARD_MAX_TOTAL_INFLATED_BYTES {
        return Err(format!(
            "Cannot read '{}': package exceeds {} bytes",
            path, HARD_MAX_TOTAL_INFLATED_BYTES
        ));
    }
    if identity(path)? != *key || data.len() as u64 != key.len {
        return Err(format!("Cannot read '{}': file changed during read", path));
    }
    Ok(data)
}

fn get<T: Any + Send + Sync>(
    path: &str,
    parse: impl FnOnce(&[u8], &Key) -> Result<T, String>,
) -> Result<Arc<T>, String> {
    get_in(cache(), path, parse)
}

fn get_in<T: Any + Send + Sync>(
    entries_lock: &Mutex<VecDeque<Entry>>,
    path: &str,
    parse: impl FnOnce(&[u8], &Key) -> Result<T, String>,
) -> Result<Arc<T>, String> {
    let key = identity(path)?;
    if let Some(found) = {
        let mut entries = entries_lock
            .lock()
            .unwrap_or_else(|poison| poison.into_inner());
        entries
            .iter()
            .position(|entry| {
                entry.key == key && entry.model.as_ref().is_some_and(|model| model.is::<T>())
            })
            .and_then(|position| entries.remove(position))
            .map(|entry| {
                let found = Arc::clone(entry.model.as_ref().expect("matched model"));
                entries.push_back(entry);
                found
            })
    } {
        return found
            .downcast::<T>()
            .map_err(|_| "cache type mismatch".to_string());
    }

    // Parse outside the lock. Concurrent misses may parse twice.
    let data = read_checked(path, &key)?;
    let model = Arc::new(parse(&data, &key)?);
    let mut entries = entries_lock
        .lock()
        .unwrap_or_else(|poison| poison.into_inner());
    let mut entry = entries
        .iter()
        .position(|entry| entry.key == key)
        .and_then(|position| entries.remove(position))
        .unwrap_or_else(|| Entry {
            key,
            model: None,
            markdown: HashMap::new(),
        });
    entry.model = Some(Arc::clone(&model) as CachedModel);
    entries.push_back(entry);
    while entries.len() > CACHE_ENTRIES {
        entries.pop_front();
    }
    Ok(model)
}

pub fn docx(path: &str) -> Result<Arc<Document>, String> {
    get(path, |bytes, _| docx_parser::parse_docx_model_native(bytes))
}

pub fn pptx(path: &str) -> Result<Arc<Presentation>, String> {
    get(path, |bytes, _| pptx_parser::parse_pptx_model_native(bytes))
}

pub struct XlsxDocument {
    pub workbook: Workbook,
    sheets: Mutex<VecDeque<(u32, Arc<Worksheet>)>>,
    key: Key,
}

pub fn xlsx(path: &str) -> Result<Arc<XlsxDocument>, String> {
    get(path, |bytes, key| {
        Ok(XlsxDocument {
            workbook: xlsx_parser::parse_workbook_model_native(bytes)?,
            sheets: Mutex::new(VecDeque::new()),
            key: key.clone(),
        })
    })
}

pub fn xlsx_sheet(
    path: &str,
    document: &XlsxDocument,
    index: u32,
    name: &str,
) -> Result<Arc<Worksheet>, String> {
    let key = identity(path)?;
    if key != document.key {
        return Err(format!("Cannot read '{}': file changed during read", path));
    }
    if let Some(sheet) = {
        let mut sheets = document
            .sheets
            .lock()
            .unwrap_or_else(|poison| poison.into_inner());
        sheets
            .iter()
            .position(|(candidate, _)| *candidate == index)
            .and_then(|position| sheets.remove(position))
            .map(|entry| {
                let sheet = Arc::clone(&entry.1);
                sheets.push_back(entry);
                sheet
            })
    } {
        return Ok(sheet);
    }
    let bytes = read_checked(path, &key)?;
    let sheet = Arc::new(xlsx_parser::parse_sheet_model_native(&bytes, index, name)?);
    let mut sheets = document
        .sheets
        .lock()
        .unwrap_or_else(|poison| poison.into_inner());
    sheets.retain(|(candidate, _)| *candidate != index);
    sheets.push_back((index, Arc::clone(&sheet)));
    while sheets.len() > CACHE_ENTRIES {
        sheets.pop_front();
    }
    Ok(sheet)
}

pub fn markdown(
    path: &str,
    kind: MarkdownKind,
    render: impl FnOnce(&[u8]) -> Result<String, String>,
) -> Result<String, String> {
    let key = identity(path)?;
    if let Some(found) = {
        let mut entries = cache().lock().unwrap_or_else(|poison| poison.into_inner());
        entries
            .iter()
            .position(|entry| entry.key == key && entry.markdown.contains_key(&kind))
            .and_then(|position| entries.remove(position))
            .map(|entry| {
                let found = Arc::clone(entry.markdown.get(&kind).expect("matched markdown"));
                entries.push_back(entry);
                found
            })
    } {
        return Ok((*found).clone());
    }
    let bytes = read_checked(path, &key)?;
    let output = render(&bytes)?;
    let mut entries = cache().lock().unwrap_or_else(|poison| poison.into_inner());
    let mut entry = entries
        .iter()
        .position(|entry| entry.key == key)
        .and_then(|position| entries.remove(position))
        .unwrap_or_else(|| Entry {
            key,
            model: None,
            markdown: HashMap::new(),
        });
    entry.markdown.insert(kind, Arc::new(output.clone()));
    entries.push_back(entry);
    while entries.len() > CACHE_ENTRIES {
        entries.pop_front();
    }
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::atomic::{AtomicUsize, Ordering};

    #[test]
    fn cache_reuses_unchanged_model_evicts_old_entries_and_checks_size() {
        let entries = Mutex::new(VecDeque::new());
        let root = std::env::temp_dir().join(format!(
            "ooxml-mcp-cache-{}-{}",
            std::process::id(),
            SystemTime::now()
                .duration_since(SystemTime::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::create_dir(&root).unwrap();
        let path = root.join("model.bin");
        fs::write(&path, b"a").unwrap();
        let path_str = path.to_str().unwrap();
        let parses = AtomicUsize::new(0);
        let parse = |bytes: &[u8], _: &Key| {
            parses.fetch_add(1, Ordering::Relaxed);
            Ok::<_, String>(bytes.to_vec())
        };
        let first = get_in(&entries, path_str, parse).unwrap();
        let reused = get_in(&entries, path_str, parse).unwrap();
        assert!(Arc::ptr_eq(&first, &reused));
        assert_eq!(parses.load(Ordering::Relaxed), 1);

        fs::write(&path, b"changed").unwrap();
        let changed = get_in(&entries, path_str, parse).unwrap();
        assert_eq!(*changed, b"changed");
        assert_eq!(parses.load(Ordering::Relaxed), 2);

        for index in 0..CACHE_ENTRIES {
            let other = root.join(format!("other-{index}.bin"));
            fs::write(&other, b"other").unwrap();
            get_in(&entries, other.to_str().unwrap(), parse).unwrap();
        }
        let reloaded = get_in(&entries, path_str, parse).unwrap();
        assert!(!Arc::ptr_eq(&changed, &reloaded));
        assert_eq!(parses.load(Ordering::Relaxed), CACHE_ENTRIES + 3);

        fs::OpenOptions::new()
            .write(true)
            .open(&path)
            .unwrap()
            .set_len(HARD_MAX_TOTAL_INFLATED_BYTES + 1)
            .unwrap();
        assert!(get_in(&entries, path_str, parse)
            .unwrap_err()
            .contains("package exceeds"));
        assert_eq!(parses.load(Ordering::Relaxed), CACHE_ENTRIES + 3);
        for index in 0..CACHE_ENTRIES {
            fs::remove_file(root.join(format!("other-{index}.bin"))).unwrap();
        }
        fs::remove_file(path).unwrap();
        fs::remove_dir(root).unwrap();
    }
}
