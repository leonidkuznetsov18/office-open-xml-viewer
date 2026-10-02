#!/usr/bin/env python3
"""Generate metadata-only reference font metrics for deterministic layout fallbacks.

The generated profiles are reference facts, not proof of the font face selected by
Canvas, the operating system, or Office. No outlines, glyph IDs, or per-glyph
advances are written to the output. Symbol coverage records cmap presence only.
OS/2 xAvgCharWidth is a scalar font metric, not a shaped text advance.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import plistlib
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from fontTools.ttLib import TTCollection, TTFont


FONT_SUFFIXES = {".otc", ".otf", ".ttc", ".ttf"}
OFFICE_ROOT = Path("/Applications/Microsoft Word.app/Contents/Resources")
MACOS_ROOT = Path("/System/Library/Fonts/Supplemental")
MACOS_PRIMARY_ROOT = Path("/System/Library/Fonts")
DATA_OUTPUT = Path("packages/core/src/fonts/reference-font-metrics-data.json")
PROVENANCE_OUTPUT = Path("scripts/reference-font-metrics-provenance.json")

# Bounded symbol domain from the #1653/#1689 slot sweeps, not all Unicode:
# Latin-1 symbols; General Punctuation; Letterlike/Number Forms, Arrows, Math,
# Misc Technical; Enclosed, Box/Block/Geometric, Misc Symbols and Dingbats.
# Presence is a font fact, independent of which language routes it to ea.
SYMBOL_COVERAGE_RANGES = ((0x00A0, 0x00FF), (0x2000, 0x206F),
                          (0x2100, 0x23FF), (0x2460, 0x27BF))


def symbol_coverage(cmap: dict[int, str] | None) -> list[int] | None:
    """Sorted inclusive endpoint pairs; None is unknown, [] is known empty."""
    if cmap is None:
        return None
    ranges: list[int] = []
    for start, end in SYMBOL_COVERAGE_RANGES:
        for code in range(start, end + 1):
            if code not in cmap or cmap[code] == ".notdef":
                continue
            if ranges and ranges[-1] == code - 1:
                ranges[-1] = code
            else:
                ranges.extend((code, code))
    return ranges


@dataclass(frozen=True)
class Source:
    id: str
    roots: tuple[tuple[str, Path], ...]
    version: str


def normalized_name(value: str) -> str:
    return " ".join(unicodedata.normalize("NFKC", value).casefold().split())


def unique_names(font: TTFont, ids: set[int]) -> list[str]:
    values: dict[str, str] = {}
    if "name" not in font:
        return []
    for record in font["name"].names:
        if record.nameID not in ids:
            continue
        try:
            value = " ".join(record.toUnicode().split())
        except (UnicodeDecodeError, AttributeError):
            continue
        key = normalized_name(value)
        if value and key not in values:
            values[key] = value
    return [values[key] for key in sorted(values)]


def preferred_name(font: TTFont, ids: Iterable[int]) -> str | None:
    if "name" not in font:
        return None
    records = font["name"].names
    for name_id in ids:
        candidates = []
        for record in records:
            if record.nameID != name_id:
                continue
            try:
                value = " ".join(record.toUnicode().split())
            except (UnicodeDecodeError, AttributeError):
                continue
            if not value:
                continue
            english = record.langID in {0, 0x0409} or record.langID == 0x8000
            windows = record.platformID == 3
            candidates.append((not english, not windows, record.langID, value))
        if candidates:
            return min(candidates)[3]
    return None


def integer(table: Any, field: str) -> int | None:
    value = getattr(table, field, None)
    return int(value) if value is not None else None


def slug(value: str) -> str:
    compact = re.sub(r"[^a-z0-9]+", "-", normalized_name(value)).strip("-")
    return compact[:48] or "unnamed"


def read_plist_version(path: Path, keys: tuple[str, ...]) -> str:
    try:
        payload = plistlib.loads(path.read_bytes())
    except (OSError, plistlib.InvalidFileException):
        return "unknown"
    return " (".join(str(payload[key]) for key in keys if payload.get(key)) + (")" if all(payload.get(key) for key in keys) and len(keys) > 1 else "")


def font_files(source: Source) -> list[tuple[str, Path]]:
    files = []
    for label, root in source.roots:
        for path in root.iterdir():
            if path.is_file() and path.suffix.lower() in FONT_SUFFIXES:
                files.append((f"{label}/{path.name}", path))
    return sorted(files, key=lambda item: (item[0].casefold(), item[0]))


def face_profile(font: TTFont, source_id: str) -> dict[str, Any] | None:
    if "head" not in font or "hhea" not in font:
        return None
    head = font["head"]
    hhea = font["hhea"]
    os2 = font["OS/2"] if "OS/2" in font else None
    family_aliases = unique_names(font, {1, 16})
    family = preferred_name(font, (16, 1)) or (family_aliases[0] if family_aliases else "Unknown")
    postscript = preferred_name(font, (6,))
    aliases_by_key = {normalized_name(value): value for value in family_aliases + unique_names(font, {6})}
    aliases = [aliases_by_key[key] for key in sorted(aliases_by_key)]
    fs_selection = integer(os2, "fsSelection") if os2 else 0
    os2_version = integer(os2, "version") if os2 else None
    italic = bool((fs_selection or 0) & 0x01 or integer(head, "macStyle") & 0x02)
    profile = {
        "source": source_id,
        "family": family,
        "aliases": aliases,
        "weight": integer(os2, "usWeightClass") if os2 else 400,
        "style": "italic" if italic else "normal",
        "unitsPerEm": integer(head, "unitsPerEm"),
        "xAvgCharWidth": integer(os2, "xAvgCharWidth") if os2 else None,
        "hhea": [integer(hhea, "ascent"), integer(hhea, "descent"), integer(hhea, "lineGap")],
    }
    provenance_code_page_range1 = (
        integer(os2, "ulCodePageRange1")
        if os2 is not None and (os2_version or 0) >= 1
        else None
    )
    # Word for Mac's controlled auto-line tests classify bits 17–20 as the
    # Far-East allocation class. Keep only that derived class at runtime;
    # absent OS/2 code-page data is unknown, not evidence of the Latin class.
    profile["farEastCodePage"] = (
        None if provenance_code_page_range1 is None
        else bool(provenance_code_page_range1 & 0x001E0000)
    )
    # Excel's DrawingML shape-text line box follows the OS/2 usWin extent
    # (issue #1604 controls: Yu Gothic, whose hhea and usWin boxes differ).
    # Null means the face has no OS/2 table.
    profile["win"] = (
        None if os2 is None
        else [integer(os2, "usWinAscent"), integer(os2, "usWinDescent")]
    )
    # PANOSE family kind and serif style (OS/2 panose bytes 1-2). PowerPoint
    # picks its application-default East Asian face from the Latin face's
    # serif class (issue #1627 controls: serif styles 2-10 took MS Mincho,
    # 11-15 MS Gothic). Null means the face has no OS/2 table.
    profile["panose"] = (
        None if os2 is None
        else [int(os2.panose.bFamilyType), int(os2.panose.bSerifStyle)]
    )
    # Whether the face's Unicode cmap maps any CJK Unified Ideograph
    # (U+4E00-U+9FFF) to a glyph. PowerPoint's empty-EA fallback (#1689
    # controls) keeps a Far-East face's own script chain only when the face
    # really draws basic CJK: SimSun-ExtB and MingLiU-ExtB declare Far-East
    # code pages but map no BMP ideograph, and fall back by PANOSE instead.
    # Null means the face has no Unicode cmap.
    cmap = font.getBestCmap() if "cmap" in font else None
    profile["cjkUnifiedIdeographs"] = (
        None if cmap is None
        else any(0x4E00 <= code <= 0x9FFF and glyph != ".notdef" for code, glyph in cmap.items())
    )
    profile["symbolCoverage"] = symbol_coverage(cmap)
    # Honor fsSelection USE_TYPO_METRICS (bit 7), like the resource parser.
    # Although introduced in OS/2 v4, installed v3 resources also set it;
    # #1689 exported symbol resources confirm their typo+gap line metrics.
    # Requiring v4 silently discards the font's declared selection.
    if os2 is not None and (fs_selection or 0) & 0x80:
        profile["typoMetrics"] = [integer(os2, "sTypoAscender"), integer(os2, "sTypoDescender"),
                                  integer(os2, "sTypoLineGap")]
    # Keep provenance identifiers stable when a source fact stops shipping in
    # the runtime profile. The identity still covers that raw OS/2 value.
    identity_profile = {
        **{key: value for key, value in profile.items()
           if key not in {"farEastCodePage", "win", "typoMetrics", "panose", "cjkUnifiedIdeographs", "symbolCoverage"}},
        "os2": None if os2 is None else {
            "codePageRange1": provenance_code_page_range1,
        },
    }
    canonical = json.dumps(identity_profile, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    profile_id = f"{source_id}:{slug(postscript or family)}:{hashlib.sha256(canonical.encode()).hexdigest()[:12]}"
    profile["_provenanceId"] = profile_id
    profile["_provenanceMetrics"] = None if os2 is None else {
        "typo": [integer(os2, "sTypoAscender"), integer(os2, "sTypoDescender"), integer(os2, "sTypoLineGap")],
        "win": [integer(os2, "usWinAscent"), integer(os2, "usWinDescent")],
        "useTypoMetrics": bool((fs_selection or 0) & 0x80),
    }
    profile["_provenanceCodePageRange1"] = provenance_code_page_range1
    return profile


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--office-root", type=Path, default=OFFICE_ROOT)
    parser.add_argument("--macos-root", type=Path, default=MACOS_ROOT)
    parser.add_argument("--macos-primary-root", type=Path, default=MACOS_PRIMARY_ROOT)
    parser.add_argument("--data-output", type=Path, default=DATA_OUTPUT)
    parser.add_argument("--provenance-output", type=Path, default=PROVENANCE_OUTPUT)
    args = parser.parse_args()

    office_version = read_plist_version(args.office_root.parent / "Info.plist", ("CFBundleShortVersionString", "CFBundleVersion"))
    system_version = read_plist_version(Path("/System/Library/CoreServices/SystemVersion.plist"), ("ProductVersion", "ProductBuildVersion"))
    sources = (
        Source("office-mac", (("DFonts", args.office_root / "DFonts"), ("OtherFonts", args.office_root / "OtherFonts")), office_version),
        Source("macos-system", (("Primary", args.macos_primary_root),), system_version),
        Source("macos-supplemental", (("Supplemental", args.macos_root),), system_version),
    )

    profiles_by_canonical: dict[str, dict[str, Any]] = {}
    provenance: list[dict[str, Any]] = []
    exclusions: list[dict[str, Any]] = []
    for source in sources:
        for relative_path, path in font_files(source):
            file_hash = hashlib.sha256(path.read_bytes()).hexdigest()
            collection = TTCollection(path, lazy=True) if path.suffix.lower() in {".ttc", ".otc"} else None
            fonts = collection.fonts if collection else [TTFont(path, lazy=True)]
            try:
                for face_index, font in enumerate(fonts):
                    if "fvar" in font:
                        exclusions.append({"source": source.id, "file": relative_path, "faceIndex": face_index, "reason": "variable-font"})
                        continue
                    profile = face_profile(font, source.id)
                    if profile is None:
                        exclusions.append({"source": source.id, "file": relative_path, "faceIndex": face_index, "reason": "missing-head-or-hhea"})
                        continue
                    profile_id = profile.pop("_provenanceId")
                    provenance_metrics = profile.pop("_provenanceMetrics")
                    provenance_code_page_range1 = profile.pop("_provenanceCodePageRange1")
                    canonical = json.dumps(profile, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
                    profiles_by_canonical.setdefault(canonical, profile)
                    provenance.append({
                        "profileId": profile_id,
                        "source": source.id,
                        "file": relative_path,
                        "faceIndex": face_index,
                        "version": preferred_name(font, (5,)),
                        "sha256": file_hash,
                        "rawOs2VerticalMetrics": provenance_metrics,
                        "rawOs2CodePageRange1": provenance_code_page_range1,
                    })
            finally:
                if collection:
                    collection.close()
                else:
                    fonts[0].close()

    source_order = {source.id: index for index, source in enumerate(sources)}
    profiles = sorted(
        profiles_by_canonical.values(),
        key=lambda item: (
            source_order[item["source"]], normalized_name(item["family"]),
            item["weight"], item["style"],
            json.dumps(item, ensure_ascii=False, sort_keys=True, separators=(",", ":")),
        ),
    )
    provenance.sort(key=lambda item: (item["source"], item["file"].casefold(), item["file"], item["faceIndex"]))
    exclusions.sort(key=lambda item: (item["source"], item["file"].casefold(), item["file"], item["faceIndex"]))
    # Intern identical repertoires across cuts and sources, including empty
    # cmaps. Runtime shares the frozen arrays; no per-glyph cache is needed.
    repertoires = sorted({tuple(p["symbolCoverage"]) for p in profiles
                          if p["symbolCoverage"] is not None})
    coverage_ids = {r: i for i, r in enumerate(repertoires)}
    for profile in profiles:
        coverage = profile["symbolCoverage"]
        profile["symbolCoverage"] = None if coverage is None else coverage_ids[tuple(coverage)]
    data = {
        "schemaVersion": 2,
        "notice": "Reference metrics are not proof of the face selected by Canvas, macOS, or Office.",
        "symbolCoverageRanges": SYMBOL_COVERAGE_RANGES,
        "symbolCoverages": repertoires,
        "profiles": profiles,
    }
    manifest = {
        "schemaVersion": 2,
        "notice": "Development provenance only; this file is not imported by the runtime metrics lookup.",
        "sources": [{"id": source.id, "version": source.version} for source in sources],
        "faces": provenance,
        "excludedFaces": exclusions,
    }
    args.data_output.parent.mkdir(parents=True, exist_ok=True)
    args.provenance_output.parent.mkdir(parents=True, exist_ok=True)
    args.data_output.write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")) + "\n")
    args.provenance_output.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")


if __name__ == "__main__":
    main()
