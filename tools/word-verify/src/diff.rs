//! Part-level comparison between the file we fed Word and what Word saved
//! after a repair. Word rewrites every XML part on save, so a line diff is
//! noise. Which parts it dropped, and which relationship types disappeared,
//! is the part that points at the cause.

use std::collections::{BTreeMap, BTreeSet};
use std::fs::File;
use std::io::Read;
use std::path::Path;

use serde::Serialize;
use zip::ZipArchive;

#[derive(Serialize, Clone, Debug, Default)]
pub struct PackageDiff {
    /// Parts present only in Word's repaired copy.
    pub added: Vec<String>,
    /// Parts Word dropped.
    pub removed: Vec<String>,
    /// Relationship types that existed before and are gone after, per source part.
    #[serde(rename = "droppedRelationships")]
    pub dropped_relationships: Vec<String>,
}

pub fn diff_docx(original: &Path, repaired: &Path) -> Result<PackageDiff, String> {
    let left = read_parts(original)?;
    let right = read_parts(repaired)?;
    let left_names: BTreeSet<_> = left.keys().cloned().collect();
    let right_names: BTreeSet<_> = right.keys().cloned().collect();
    let mut dropped = Vec::new();
    for (name, bytes) in &left {
        if !name.ends_with(".rels") {
            continue;
        }
        let before = rel_types(bytes);
        let after = right.get(name).map(|b| rel_types(b)).unwrap_or_default();
        for ty in before.difference(&after) {
            dropped.push(format!("{name}: {ty}"));
        }
    }
    Ok(PackageDiff {
        added: right_names.difference(&left_names).cloned().collect(),
        removed: left_names.difference(&right_names).cloned().collect(),
        dropped_relationships: dropped,
    })
}

fn rel_types(bytes: &[u8]) -> BTreeSet<String> {
    let text = String::from_utf8_lossy(bytes);
    let mut out = BTreeSet::new();
    for chunk in text.split("Type=\"").skip(1) {
        if let Some(end) = chunk.find('"') {
            let ty = &chunk[..end];
            out.insert(ty.rsplit('/').next().unwrap_or(ty).to_string());
        }
    }
    out
}

fn read_parts(path: &Path) -> Result<BTreeMap<String, Vec<u8>>, String> {
    let file = File::open(path).map_err(|err| format!("open {}: {err}", path.display()))?;
    let mut zip = ZipArchive::new(file).map_err(|err| format!("zip {}: {err}", path.display()))?;
    let mut parts = BTreeMap::new();
    for i in 0..zip.len() {
        let mut entry = zip.by_index(i).map_err(|err| err.to_string())?;
        if !entry.is_file() {
            continue;
        }
        let name = entry.name().replace('\\', "/");
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).map_err(|err| err.to_string())?;
        parts.insert(name, bytes);
    }
    Ok(parts)
}
