#!/usr/bin/env python3
"""
Compare two EPFL study plan Excel workbooks and highlight the differences.

The script reads the "old" workbook, finds the matching sheets in the "new"
workbook, and annotates the old workbook in-place by highlighting every cell
that changed. A summary sheet called "Diff Summary" is also added to the output
workbook so that additions/removals are easy to review.

Example:
    python compare_study_plans.py \
        --old "1 - Excel SAC 2022/SC-Plan_d_etudes-2022-2023.xlsx" \
        --new "3 - Excel SAC 2023/SC-Plan_d_etudes-2023-2024.xlsx" \
        --output "SC-Plan_d_etudes-2022-2023_with_diffs.xlsx"
"""

from __future__ import annotations

import argparse
import difflib
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import MergedCell

# --- Configuration -----------------------------------------------------------------

# All the header labels we want to recognize. Values are lists of keywords that must
# appear (after accent stripping/lower-casing) for the column to be identified.
FIELD_PATTERNS: Dict[str, Sequence[str]] = {
    "code": ("code", "codes"),
    "course": ("matiere", "matieres", "subjects", "course"),
    # Make the branch type detection strict so we do not mis-detect "Type examen".
    "type": ("type de branche", "type de branches", "branches"),
    "teacher": ("enseignant", "enseignants", "professeur", "teacher", "teachers"),
    "sections": ("section", "sections"),
    "credits": ("credit", "credits", "ects"),
    "semesters": ("semestre", "semestres"),
    "period": ("periode", "period"),
    "language": ("langue",),
    "specializations": ("specialisation", "specialisations"),
    "coeff": ("coeff", "coefficient"),
}

# Determines how many columns should be captured for each field. Fields with
# mode == "gap" expand until the next identified header but no further than
# the provided max width.
FIELD_SPAN_RULES: Dict[str, Dict[str, int]] = {
    "period": {"mode": "fixed", "width": 2},
    "specializations": {"mode": "gap", "max": 12},
    "semesters": {"mode": "gap", "max": 8},
}

# Colors used for highlights.
MODIFIED_FILL = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")
REMOVED_FILL = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
ADDED_FILL = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")

# Fields intentionally excluded from value comparison.
EXCLUDED_COMPARE_FIELDS = {"teacher", "period", "language", "semesters"}


# --- Helper data structures ---------------------------------------------------------

@dataclass
class SheetLayout:
    """Column ranges (start/end) for each field inside a sheet section."""

    field_ranges: Dict[str, Tuple[int, int]]
    header_labels: Dict[str, str]


@dataclass
class CourseRow:
    """Normalized view of a course row inside a worksheet."""

    sheet: str
    row_idx: int
    display_code: str
    code_key: str
    base_code_key: str
    layout_ranges: Dict[str, Tuple[int, int]]
    field_values: Dict[str, List[Optional[object]]]
    header_labels: Dict[str, str]
    group_key: str
    group_label: str
    group_order: int
    compare_fields: Sequence[str]
    used: bool = field(default=False)


# --- Text utilities -----------------------------------------------------------------

def strip_accents(value: str) -> str:
    """Remove accents/diacritics from a string."""

    normalized = unicodedata.normalize("NFKD", value)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def normalize_text(value: Optional[object]) -> str:
    """Normalize a header cell so that pattern matching becomes easier."""

    if value is None:
        return ""
    if not isinstance(value, str):
        value = str(value)
    value = value.replace("\n", " ")
    value = strip_accents(value).lower()
    return re.sub(r"\s+", " ", value).strip()


def format_display(value: Optional[object]) -> str:
    """Create a nice human-readable representation for summary/comment text."""

    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def normalize_value(value: Optional[object]) -> str:
    """Normalize cell values before comparison."""

    text = format_display(value)
    return text.strip()


def normalize_group_label(text: str) -> str:
    """Normalize group/bloc labels so that similar wording still matches."""

    if not text:
        return ""
    normalized = normalize_text(text)
    normalized = re.sub(r"[^a-z0-9]+", " ", normalized)
    return normalized.strip()


def group_kind(group_key: str) -> str:
    """Return the high-level type of a grouping label."""

    if group_key.startswith("bloc"):
        return "bloc"
    if group_key.startswith("block"):
        return "bloc"
    if group_key.startswith("groupe"):
        return "groupe"
    if group_key.startswith("group"):
        return "groupe"
    if group_key.startswith("specialisation"):
        return "specialisation"
    if group_key.startswith("specialization"):
        return "specialisation"
    if group_key.startswith("option"):
        return "option"
    if group_key.startswith("options"):
        return "option"
    if group_key.startswith("gap-"):
        return "gap"
    return "other"


def looks_like_group_header(text: str) -> bool:
    """True only for actual bloc/group/specialization title rows."""

    normalized = normalize_text(text)
    return bool(
        re.match(
            r"^(bloc|block|groupe|group|specialisation|specialization|option|options)\b",
            normalized,
        )
    )


CODE_KEY_PATTERN = re.compile(r"[^0-9A-Za-z]+")


def canonical_code(value: Optional[object]) -> Tuple[str, str, str]:
    """Return original display string plus normalized keys for matching."""

    if value is None:
        return "", "", ""
    text = format_display(value)
    text = text.replace("–", "-").replace("—", "-")
    text = text.replace("\xa0", " ").strip()
    base = text.split("(")[0].strip()
    key = CODE_KEY_PATTERN.sub("", text).upper()
    base_key = CODE_KEY_PATTERN.sub("", base).upper()
    return text, key, base_key


def looks_like_course_code(code: str) -> bool:
    """Heuristic to decide if a given string is a course code."""

    if not code:
        return False
    lowered = code.lower().strip()
    # Course codes never contain spaces; this filters out note/comment rows.
    if re.search(r"\s", lowered):
        return False
    if lowered.startswith("total") or lowered.startswith("totaux"):
        return False
    if lowered.startswith("bloc"):
        return False
    return bool(re.search(r"\d", code))


# --- Header/layout detection --------------------------------------------------------

def detect_header(
    row_values: Sequence[Optional[object]],
) -> Optional[Tuple[Dict[str, int], Dict[str, str]]]:
    """Return a mapping from field name to column index plus header labels."""

    mapping: Dict[str, int] = {}
    labels: Dict[str, str] = {}
    for idx, value in enumerate(row_values, start=1):
        normalized = normalize_text(value)
        if not normalized:
            continue
        for field, keywords in FIELD_PATTERNS.items():
            if field in mapping:
                continue
            if field == "code":
                # Avoid false positives like "codesign" in course names while
                # accepting both "Code" and "Codes" headers.
                matched = bool(re.search(r"\bcodes?\b", normalized))
            elif field == "credits":
                # Prevent accidental match of "SUBJECTS" (contains "...ects").
                matched = bool(re.search(r"\bcredits?\b|\bects\b", normalized))
            else:
                matched = any(keyword in normalized for keyword in keywords)
            if matched:
                mapping[field] = idx
                labels[field] = normalized
    # A real header row should identify code plus at least one other known field.
    if "code" in mapping and len(mapping) >= 2:
        return mapping, labels
    return None


def compute_field_ranges(
    header_positions: Dict[str, int], max_column: int
) -> Dict[str, Tuple[int, int]]:
    """Expand header positions into column ranges."""

    sorted_fields = sorted(header_positions.items(), key=lambda item: item[1])
    ranges: Dict[str, Tuple[int, int]] = {}
    for index, (field, start_col) in enumerate(sorted_fields):
        next_start = (
            sorted_fields[index + 1][1] if index + 1 < len(sorted_fields) else max_column + 1
        )
        rule = FIELD_SPAN_RULES.get(field, {"mode": "fixed", "width": 1})
        if rule["mode"] == "fixed":
            end_col = start_col + rule.get("width", 1) - 1
        else:  # "gap"
            max_width = rule.get("max", 1)
            gap_end = next_start - 1 if next_start > start_col else max_column
            end_col = min(gap_end, start_col + max_width - 1)
        end_col = min(end_col, max_column)
        if end_col < start_col:
            end_col = start_col
        ranges[field] = (start_col, end_col)
    return ranges


def slice_row_values(
    row_values: Sequence[Optional[object]], start: int, end: int
) -> List[Optional[object]]:
    """Return the values from start..end (1-based coordinates)."""

    result: List[Optional[object]] = []
    for col in range(start, end + 1):
        idx = col - 1
        result.append(row_values[idx] if idx < len(row_values) else None)
    return result


def parse_sheet(ws: Worksheet) -> List[CourseRow]:
    """Extract all course rows from a worksheet."""

    courses: List[CourseRow] = []
    current_layout: Optional[SheetLayout] = None
    current_group_key: str = ""
    current_group_label: str = ""
    group_counter = 0
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        row_values = list(row)
        header_detection = detect_header(row_values)
        if header_detection:
            header_map, header_labels = header_detection
            field_ranges = compute_field_ranges(header_map, ws.max_column)
            current_layout = SheetLayout(field_ranges=field_ranges, header_labels=header_labels)
            current_group_key = ""
            current_group_label = ""
            continue
        if not current_layout or "code" not in current_layout.field_ranges:
            continue
        code_start, code_end = current_layout.field_ranges["code"]
        code_cell = slice_row_values(row_values, code_start, code_end)[0]
        course_cell = None
        if "course" in current_layout.field_ranges:
            c_start, c_end = current_layout.field_ranges["course"]
            course_cell = slice_row_values(row_values, c_start, c_end)[0]
        display_code, code_key, base_code = canonical_code(code_cell)
        # Update grouping markers (Bloc/Groupe/Spécialisation/etc.).
        # Some workbooks place group labels in "Matières", others in "Code".
        course_text = normalize_value(course_cell)
        code_text = normalize_value(code_cell)
        marker_text = course_text or code_text
        normalized_marker_text = normalize_text(marker_text)
        if marker_text and not looks_like_course_code(marker_text) and looks_like_group_header(marker_text):
            group_counter += 1
            current_group_key = normalize_group_label(marker_text)
            current_group_label = marker_text
            continue
        # Treat a fully blank separator row as a new anonymous group boundary.
        if not display_code and not course_text:
            group_counter += 1
            current_group_key = f"gap-{group_counter}"
            current_group_label = ""
            continue
        if not code_key or not looks_like_course_code(display_code):
            continue
        # Capture all fields present in the detected header (besides code); keep them
        # for comparison. Because the user aligned the tables, we can compare every
        # non-code column we detect.
        fields_to_capture = set(current_layout.field_ranges.keys())
        field_ranges = {
            field: rng
            for field, rng in current_layout.field_ranges.items()
            if field in fields_to_capture
        }
        field_values = {
            field: slice_row_values(row_values, start, end)
            for field, (start, end) in field_ranges.items()
        }
        compare_fields = [
            f
            for f in field_ranges.keys()
            if f != "code" and f not in EXCLUDED_COMPARE_FIELDS
        ]
        courses.append(
            CourseRow(
                sheet=ws.title,
                row_idx=row_idx,
                display_code=display_code,
                code_key=code_key,
                base_code_key=base_code,
                layout_ranges=field_ranges,
                field_values=field_values,
                header_labels=current_layout.header_labels,
                group_key=current_group_key,
                group_label=current_group_label,
                group_order=group_counter,
                compare_fields=compare_fields,
            )
        )
    return courses


# --- Sheet comparison ----------------------------------------------------------------

def select_best_match(old_row: CourseRow, candidates: List[CourseRow]) -> int:
    """Pick the most similar candidate row."""

    best_idx = 0
    best_score = -1
    for idx, candidate in enumerate(candidates):
        if candidate.used:
            continue
        score = 0
        fields = [f for f in old_row.compare_fields if f in candidate.compare_fields]
        for field in fields:
            old_vals = old_row.field_values.get(field)
            new_vals = candidate.field_values.get(field)
            if not old_vals or not new_vals:
                continue
            normalized_old = [normalize_value(val) for val in old_vals]
            normalized_new = [normalize_value(val) for val in new_vals]
            score += sum(1 for a, b in zip(normalized_old, normalized_new) if a == b)
        if score > best_score:
            best_idx = idx
            best_score = score
    return best_idx


def find_name_match(old_row: CourseRow, candidates: List[CourseRow], threshold: float = 0.7) -> Optional[CourseRow]:
    """Fallback: match by course name similarity when codes don't align."""

    def course_name(row: CourseRow) -> str:
        vals = row.field_values.get("course", [])
        return normalize_text(vals[0]) if vals else ""

    old_name = course_name(old_row)
    best = None
    best_score = threshold
    for cand in candidates:
        score = difflib.SequenceMatcher(None, old_name, course_name(cand)).ratio()
        if score > best_score:
            best_score = score
            best = cand
    return best


def annotate_cell(cell, fill: PatternFill, message: str) -> None:
    """Apply highlighting/comment to a cell."""

    # If this is part of a merged cell range, apply formatting/comment to the anchor.
    if isinstance(cell, MergedCell):
        ws = cell.parent
        anchor = None
        for merged in ws.merged_cells.ranges:
            if (
                merged.min_row <= cell.row <= merged.max_row
                and merged.min_col <= cell.column <= merged.max_col
            ):
                anchor = ws.cell(row=merged.min_row, column=merged.min_col)
                break
        if anchor is None:
            return
        cell = anchor
    cell.fill = fill
    if message:
        if cell.comment:
            text = f"{cell.comment.text}\n{message}"
        else:
            text = message
        cell.comment = Comment(text, "Diff")


def insert_new_row(
    target_ws: Worksheet,
    template_row: CourseRow,
    target_layout_ranges: Dict[str, Tuple[int, int]],
    diff_log: List[Dict[str, str]],
    note: str,
    insert_at: Optional[int] = None,
) -> None:
    """Insert a new row into the target worksheet using data from template_row."""

    insert_at = insert_at or min(template_row.row_idx, target_ws.max_row + 1)
    target_ws.insert_rows(insert_at)
    for field, (start, end) in target_layout_ranges.items():
        values = template_row.field_values.get(field, [])
        width = end - start + 1
        for offset in range(width):
            value = values[offset] if offset < len(values) else None
            col = start + offset
            cell = target_ws.cell(row=insert_at, column=col)
            cell.value = value
            annotate_cell(cell, ADDED_FILL, note)
    template_row.used = True
    diff_log.append(
        {
            "sheet": template_row.sheet,
            "code": template_row.display_code,
            "field": "row",
            "old": "",
            "new": "",
            "details": note,
        }
    )


def compare_rows(
    ws: Worksheet,
    old_row: CourseRow,
    new_row: CourseRow,
    diff_log: List[Dict[str, str]],
    row_offset: int,
) -> int:
    """Highlight per-field differences between two course rows."""

    fields = [f for f in old_row.compare_fields if f in new_row.compare_fields]
    for field in fields:
        if field not in old_row.field_values or field not in new_row.field_values:
            continue
        old_values = old_row.field_values[field]
        new_values = new_row.field_values[field]
        start_col, _ = old_row.layout_ranges[field]
        max_len = max(len(old_values), len(new_values))
        for offset in range(max_len):
            old_val = old_values[offset] if offset < len(old_values) else None
            new_val = new_values[offset] if offset < len(new_values) else None
            if normalize_value(old_val) == normalize_value(new_val):
                continue
            column = start_col + offset
            cell = ws.cell(row=old_row.row_idx + row_offset, column=column)
            annotation = f"2023-2024 value: {format_display(new_val)}"
            annotate_cell(cell, MODIFIED_FILL, annotation)
            diff_log.append(
                {
                    "sheet": old_row.sheet,
                    "code": old_row.display_code,
                    "field": field,
                    "old": format_display(old_val),
                    "new": format_display(new_val),
                    "details": "value changed",
                }
            )
    return 0


def mark_removed_row(
    ws: Worksheet, row: CourseRow, diff_log: List[Dict[str, str]], row_offset: int
) -> None:
    """Highlight an entire course row that disappeared in the new workbook."""

    for field, (start, end) in row.layout_ranges.items():
        for col in range(start, end + 1):
            cell = ws.cell(row=row.row_idx + row_offset, column=col)
            annotate_cell(cell, REMOVED_FILL, "No longer present in 2023-2024")
    diff_log.append(
        {
            "sheet": row.sheet,
            "code": row.display_code,
            "field": "row",
            "old": "",
            "new": "",
            "details": "course removed in 2023-2024 workbook",
        }
    )


def compare_sheet(
    old_ws: Worksheet,
    old_rows: List[CourseRow],
    new_rows: List[CourseRow],
    diff_log: List[Dict[str, str]],
) -> None:
    """Compare every course row in a sheet."""

    if not old_rows:
        return
    global_code_index = defaultdict(list)
    global_base_index = defaultdict(list)
    for row in new_rows:
        global_code_index[row.code_key].append(row)
        global_base_index[row.base_code_key].append(row)
    # Group rows by bloc/groupe to ensure adds stay within their block.
    def group_rows(rows: List[CourseRow]):
        grouped = defaultdict(list)
        order = []
        for r in sorted(rows, key=lambda r: (r.group_order, r.row_idx)):
            key = r.group_key or ""
            grouped[key].append(r)
            if key not in order:
                order.append(key)
        return grouped, order

    grouped_old, order_old = group_rows(old_rows)
    grouped_new, order_new = group_rows(new_rows)
    default_old_layout_ranges = old_rows[0].layout_ranges

    def map_group_keys(old_keys: List[str], new_keys: List[str]) -> Tuple[Dict[str, Optional[str]], List[str]]:
        remaining = list(new_keys)
        mapping: Dict[str, Optional[str]] = {}
        for ok in old_keys:
            if ok in remaining:
                mapping[ok] = ok
                remaining = [k for k in remaining if k != ok]
                continue
            # Synthetic gap groups should not be fuzzy-matched by label.
            if group_kind(ok) == "gap":
                mapping[ok] = None
                continue
            best_key = None
            best_score = 0.0
            ok_kind = group_kind(ok)
            for nk in list(remaining):
                nk_kind = group_kind(nk)
                # Keep matching within the same grouping family.
                if ok_kind != "gap" and nk_kind != "gap" and ok_kind != nk_kind:
                    continue
                score = difflib.SequenceMatcher(None, ok, nk).ratio()
                if score > best_score:
                    best_score = score
                    best_key = nk
            if best_key and best_score >= 0.6:
                mapping[ok] = best_key
                remaining = [k for k in remaining if k != best_key]
            else:
                mapping[ok] = None
        # If labels changed a lot, map remaining bloc/group sections by course-code overlap.
        for ok in old_keys:
            if mapping.get(ok) is not None:
                continue
            ok_kind = group_kind(ok)
            old_codes = {r.base_code_key for r in grouped_old.get(ok, []) if r.base_code_key}
            if not old_codes:
                continue
            best_key = None
            best_overlap = 0
            for nk in list(remaining):
                nk_kind = group_kind(nk)
                if ok_kind != "gap" and nk_kind != "gap" and ok_kind != nk_kind:
                    continue
                new_codes = {r.base_code_key for r in grouped_new.get(nk, []) if r.base_code_key}
                overlap = len(old_codes & new_codes)
                if overlap > best_overlap:
                    best_overlap = overlap
                    best_key = nk
            if best_key and best_overlap > 0:
                mapping[ok] = best_key
                remaining = [k for k in remaining if k != best_key]
        # Fallback: map any still-unmatched old groups to remaining new groups by order.
        remaining_order = list(remaining)
        for ok in old_keys:
            # Only force positional fallback for synthetic gap groups.
            if mapping.get(ok) is None and group_kind(ok) == "gap" and remaining_order:
                chosen = remaining_order.pop(0)
                mapping[ok] = chosen
                remaining = [k for k in remaining if k != chosen]
        return mapping, remaining

    group_mapping, unmatched_new_groups = map_group_keys(order_old, order_new)

    cumulative_offset = 0

    for group_key in order_old:
        old_group_rows = grouped_old.get(group_key, [])
        mapped_new_key = group_mapping.get(group_key)
        new_group_rows = grouped_new.get(mapped_new_key, []) if mapped_new_key else []

        # Reset usage per bloc to avoid cross-bloc depletion.
        for r in new_group_rows:
            r.used = False

        code_index = defaultdict(list)
        base_index = defaultdict(list)
        for row in new_group_rows:
            code_index[row.code_key].append(row)
            base_index[row.base_code_key].append(row)

        group_start = min(r.row_idx for r in old_group_rows) if old_group_rows else old_ws.max_row
        group_end = max(r.row_idx for r in old_group_rows) if old_group_rows else group_start
        inserted_here = 0

        for old_row in old_group_rows:
            current_offset = cumulative_offset + inserted_here
            candidates = [row for row in code_index.get(old_row.code_key, []) if not row.used]
            if not candidates:
                candidates = [
                    row for row in base_index.get(old_row.base_code_key, []) if not row.used
                ]
            if not candidates:
                name_candidate = find_name_match(
                    old_row, [r for r in new_group_rows if not r.used]
                )
                if name_candidate:
                    candidates = [name_candidate]
            if not candidates:
                # No match inside this bloc: consider it removed here.
                mark_removed_row(old_ws, old_row, diff_log, current_offset)
                continue
            best_idx = select_best_match(old_row, candidates)
            match = candidates[best_idx]
            match.used = True
            inserted_delta = compare_rows(old_ws, old_row, match, diff_log, current_offset)
            inserted_here += inserted_delta

        for row in new_group_rows:
            if not row.used:
                insert_pos = group_end + cumulative_offset + inserted_here + 1
                target_layout_ranges = (
                    old_group_rows[0].layout_ranges if old_group_rows else default_old_layout_ranges
                )
                insert_new_row(
                    old_ws,
                    row,
                    target_layout_ranges,
                    diff_log,
                    "New course in 2023-2024 workbook",
                    insert_at=insert_pos,
                )
                inserted_here += 1

        cumulative_offset += inserted_here

    # Handle entirely new groups (if any) by appending at the end in their order.
    for group_key in unmatched_new_groups:
        new_group_rows = grouped_new.get(group_key, [])
        for row in new_group_rows:
            insert_pos = old_ws.max_row + 1
            insert_new_row(
                old_ws,
                row,
                default_old_layout_ranges,
                diff_log,
                f"New course in 2023-2024 workbook (new bloc {row.group_label or ''})",
                insert_at=insert_pos,
            )

    # Safety net: if any new rows remain unused, append them at the end.
    for row in new_rows:
        if not row.used:
            insert_new_row(
                old_ws,
                row,
                default_old_layout_ranges,
                diff_log,
                f"New course in 2023-2024 workbook (unmatched bloc {row.group_label or ''})",
                insert_at=old_ws.max_row + 1,
            )


# --- Workbook level orchestration ---------------------------------------------------

def parse_workbook(wb: Workbook) -> Dict[str, List[CourseRow]]:
    """Parse every worksheet in a workbook."""

    data: Dict[str, List[CourseRow]] = {}
    for name in wb.sheetnames:
        ws = wb[name]
        data[name] = parse_sheet(ws)
    return data


def normalize_sheet_name(name: str) -> str:
    """Normalize sheet names for fuzzy matching."""

    text = strip_accents(name).lower()
    return re.sub(r"[^a-z0-9]", "", text)


def match_sheets(old_names: Sequence[str], new_names: Sequence[str]) -> Dict[str, Optional[str]]:
    """Map each old sheet name to the closest sheet name in the new workbook."""

    matches: Dict[str, Optional[str]] = {}
    available = set(new_names)
    normalized_new = {name: normalize_sheet_name(name) for name in new_names}
    for old in old_names:
        norm_old = normalize_sheet_name(old)
        exact = [name for name, norm in normalized_new.items() if norm == norm_old and name in available]
        if exact:
            chosen = exact[0]
            matches[old] = chosen
            available.remove(chosen)
            continue
        best_name = None
        best_score = 0.0
        for candidate in available:
            score = difflib.SequenceMatcher(
                None, norm_old, normalized_new[candidate]
            ).ratio()
            if score > best_score:
                best_score = score
                best_name = candidate
        matches[old] = best_name if best_score >= 0.5 else None
        if best_name:
            available.remove(best_name)
    return matches


def write_summary_sheet(wb: Workbook, diff_log: List[Dict[str, str]]) -> None:
    """Create (or replace) the summary sheet containing all the differences."""

    if "Diff Summary" in wb.sheetnames:
        del wb["Diff Summary"]
    ws = wb.create_sheet("Diff Summary")
    ws.append(["Sheet", "Code", "Field", "Old Value", "New Value", "Details"])
    for entry in diff_log:
        ws.append(
            [
                entry.get("sheet", ""),
                entry.get("code", ""),
                entry.get("field", ""),
                entry.get("old", ""),
                entry.get("new", ""),
                entry.get("details", ""),
            ]
        )


def run(old_path: str, new_path: str, output_path: str) -> None:
    """Main orchestration logic."""

    old_wb = load_workbook(old_path)
    new_wb = load_workbook(new_path, data_only=True)
    old_data = parse_workbook(old_wb)
    new_data = parse_workbook(new_wb)
    diff_log: List[Dict[str, str]] = []

    sheet_mapping = match_sheets(old_wb.sheetnames, new_wb.sheetnames)
    for old_name, new_name in sheet_mapping.items():
        if new_name is None:
            continue
        compare_sheet(old_wb[old_name], old_data.get(old_name, []), new_data.get(new_name, []), diff_log)

    write_summary_sheet(old_wb, diff_log)
    old_wb.save(output_path)


# --- CLI ----------------------------------------------------------------------------

def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Highlight differences between two study plan Excel workbooks.")
    parser.add_argument("--old", required=True, help="Path to the reference workbook (e.g. 2022-2023).")
    parser.add_argument("--new", required=True, help="Path to the updated workbook (e.g. 2023-2024).")
    parser.add_argument(
        "--output",
        required=True,
        help="Path of the Excel file that will receive the highlighted changes.",
    )
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    run(args.old, args.new, args.output)


if __name__ == "__main__":
    main()
