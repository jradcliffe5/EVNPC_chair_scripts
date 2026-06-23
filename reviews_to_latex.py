#!/usr/bin/env python3
"""Build a LaTeX summary document from reviewer text files.

The input review files are expected to live in a directory (default: ./reviews)
and to follow the structure produced by the EVN chair template, i.e. each file
contains blocks of proposals separated by a 100-character line of '=' symbols.
Only proposals with at least one non-empty field (grade, referee comments,
technical review, or time recommended) are included in the summary.
"""

from __future__ import annotations

import argparse
import csv
import io
import math
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from textwrap import dedent
from typing import Dict, Iterable, List, Optional, Sequence

SEPARATOR = "=" * 100
FIELD_NAMES = [
    "Grade",
    "Referee comments",
    "Technical review",
    "Time recommended",
]
SIMPLE_FIELD_NAMES = [
    "Grade",
    "General remark",
    "Strengths",
    "Weaknesses",
    "Referee comments",
    "Technical review",
    "Time recommended",
]
SIMPLE_FIELD_ALIASES = {
    "grade": "Grade",
    "general remark": "General remark",
    "general remarks": "General remark",
    "strengths": "Strengths",
    "weaknesses": "Weaknesses",
    "referee comments": "Referee comments",
    "technical review": "Technical review",
    "time recommended": "Time recommended",
}
PROPOSAL_CODE_RE = re.compile(r"^[A-Z]\d{2}[A-Z]\d{3}$", re.IGNORECASE)
SPECIAL_GRADE_RE = re.compile(r"\b(resubmission|resubmit|reject)\b", re.IGNORECASE)


@dataclass
class ReviewEntry:
    reviewer: str
    source_file: Path
    grade: str = ""
    referee_comments: str = ""
    technical_review: str = ""
    time_recommended: str = ""
    role: Optional[int] = None  # 1 for first reviewer, 2 for second reviewer


@dataclass
class ProposalSummary:
    code: str
    title: str
    pi: str = ""
    networks: str = ""
    wavelengths: str = ""
    reviews: List[ReviewEntry] = field(default_factory=list)


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert individual review text files into a LaTeX summary document.",
        epilog=dedent(
            """\
            Examples:
              reviews_to_latex.py -r reviews -o summary.tex
              reviews_to_latex.py -r reviews -a reviewer_assignments.txt -t "EVN Reviews" -V "Draft 3"
            """
        ),
    )
    parser.add_argument(
        "-r",
        "--reviews-dir",
        type=Path,
        default=Path("reviews"),
        help="Directory containing renamed review text files (default: ./reviews).",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("review_summary.tex"),
        help="Path for the generated LaTeX file (default: ./review_summary.tex).",
    )
    parser.add_argument(
        "-t",
        "--title",
        default="EVN Review Summary",
        help="Title for the LaTeX document.",
    )
    parser.add_argument(
        "-V",
        "--version",
        default="",
        help="Optional version string to display beneath the title.",
    )
    parser.add_argument(
        "-a",
        "--assignments",
        type=Path,
        default=None,
        help="Optional reviewer assignment summary file to tag first/second reviewers.",
    )
    parser.add_argument(
        "--agenda-txt",
        type=Path,
        default=None,
        help=(
            "Optional path for agenda-friendly plain-text output "
            "(proposal/PI/title + reviewer initials + pre-grade + std + N>2.0 + conflict)."
        ),
    )
    parser.add_argument(
        "--code-mapping",
        type=Path,
        default=None,
        help=(
            "Optional file mapping EVN session codes to proposal codes "
            "(e.g. E26A001 -> EC107).  When provided, section titles are "
            "rendered as 'E26A001/EC107: Title'."
        ),
    )
    return parser.parse_args(argv)


UNICODE_REPLACEMENTS = {
    " ": " ",
    " ": " ",
    "‐": "-",
    "‑": "-",
    "‒": "-",
    "–": "--",
    "—": "---",
    "−": "-",
    "…": "...",
    "‘": "'",
    "’": "'",
    "“": "``",
    "”": "''",
    "″": r"$''$",
    "µ": r"$\mu$",
    "°": r"$^{\circ}$",
    "×": r"$\times$",
    "∼": r"$\sim$",
    "≤": r"$\le$",
    "≥": r"$\ge$",
    "α": r"$\alpha$",
    "β": r"$\beta$",
    "γ": r"$\gamma$",
    "π": r"$\pi$",
    "μ": r"$\mu$",
}

# Unicode characters that should remain in the generated .tex but need LaTeX
# declarations to compile under pdfLaTeX.
UNICODE_LATEX_DECLARATIONS = {
    "00B1": r"\ensuremath{\pm}",
    "0144": r"\'{n}",
    "03C3": r"\ensuremath{\sigma}",
    "2192": r"\ensuremath{\rightarrow}",
    "2248": r"\ensuremath{\approx}",
    "FF5E": r"\textasciitilde{}",
}


def normalize_unicode(text: str) -> str:
    """Normalize common Unicode punctuation/symbols while preserving UTF-8 text."""
    if not text:
        return text
    text = unicodedata.normalize("NFC", text)
    for source, replacement in UNICODE_REPLACEMENTS.items():
        text = text.replace(source, replacement)
    return text


def normalize_paragraphs(text: str) -> str:
    """Collapse line breaks inside paragraphs; keep blank-line paragraph breaks."""
    if not text:
        return text
    paragraphs = re.split(r"\n\s*\n", text)
    cleaned_paragraphs = []
    for para in paragraphs:
        # Join any wrapped lines within the paragraph.
        pieces = [line.strip() for line in para.splitlines() if line.strip()]
        if not pieces:
            continue
        cleaned_paragraphs.append(" ".join(pieces))
    return "\n\n".join(cleaned_paragraphs)


def latex_escape(text: str) -> str:
    """Escape LaTeX special characters in the supplied text."""
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    pattern = re.compile("|".join(re.escape(key) for key in replacements))
    escaped = pattern.sub(lambda match: replacements[match.group()], text)
    return normalize_unicode(escaped)

def reviewer_initials(name: str) -> str:
    """Return uppercase initials derived from the reviewer name."""
    tokens = [token for token in name.replace("_", " ").split() if token]
    initials = "".join(token[0].upper() for token in tokens if token[0].isalpha())
    return initials or "?"


def parse_numeric_grade(value: str) -> Optional[float]:
    """Extract the first numeric token from the grade string."""
    if SPECIAL_GRADE_RE.search(value or ""):
        return 4.0
    match = re.search(r"-?\d+(?:\.\d+)?", value)
    if not match:
        return None
    try:
        return float(match.group())
    except ValueError:
        return None


def normalise_person_name(value: str) -> str:
    """Canonical lowercase name for matching assignments."""
    return " ".join(value.lower().split())


def normalise_reviewer_from_filename(path: Path) -> str:
    """Infer the reviewer name from the filename `PREFIX_Name_Surname.ext`."""
    stem = path.stem
    parts = stem.split("_", 1)
    if len(parts) == 2:
        _, reviewer = parts
        return reviewer.replace("_", " ").strip() or "Unknown reviewer"
    return "Unknown reviewer"


def role_superscript(role: Optional[int]) -> str:
    """Return latex superscript for reviewer role."""
    if role == 1:
        return r"\textsuperscript{1}"
    if role == 2:
        return r"\textsuperscript{2}"
    return ""


def role_sort_key(role: Optional[int]) -> int:
    """Sort order: primary (1), secondary (2), then unassigned."""
    if role == 1:
        return 0
    if role == 2:
        return 1
    return 2


def read_text_with_fallback(path: Path) -> str:
    """Read text as UTF-8, falling back to cp1252 for legacy files."""
    try:
        # utf-8-sig strips a leading BOM (which str.strip() leaves in place and
        # would otherwise drop the first proposal during parsing).
        return path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as utf8_error:
        try:
            text = path.read_text(encoding="cp1252")
        except UnicodeDecodeError:
            raise utf8_error
        print(
            f"Warning: decoded non-UTF-8 file using cp1252: {path}",
            file=sys.stderr,
        )
        return text


def load_assignments(path: Path) -> Dict[str, Dict[str, int]]:
    """Parse reviewer assignment summary to map proposal codes to reviewer roles."""
    mapping: Dict[str, Dict[str, int]] = {}
    data = read_text_with_fallback(path)
    if not data.strip():
        return mapping

    lines = data.splitlines()
    if not lines:
        return mapping

    header_line = lines[0].strip().lower()
    if "," in header_line and "proposal" in header_line:
        reader = csv.DictReader(io.StringIO(data))
        for row in reader:
            row_lower = {key.lower(): (value or "").strip() for key, value in row.items()}
            code = row_lower.get("proposal", "")
            if not code:
                continue
            code_map = mapping.setdefault(code, {})
            first = row_lower.get("first reviewer", "") or row_lower.get(
                "primary reviewer", ""
            )
            second = row_lower.get("second reviewer", "") or row_lower.get(
                "secondary reviewer", ""
            )
            if first:
                code_map[normalise_person_name(first)] = 1
            if second:
                code_map[normalise_person_name(second)] = 2
        return mapping

    pattern = re.compile(r"([A-Z]\d{2}[A-Z]\d{3})\s*\(([^)]+)\)", re.IGNORECASE)
    for line in lines:
        if ":" not in line:
            continue
        name_part, assignments_part = line.split(":", 1)
        reviewer_name = normalise_person_name(name_part.strip())
        if not reviewer_name:
            continue
        for entry in assignments_part.split(","):
            entry = entry.strip()
            if not entry:
                continue
            match = pattern.search(entry)
            if not match:
                continue
            code = match.group(1).strip()
            role_label = match.group(2).strip().lower()
            if "first" in role_label:
                role = 1
            elif "second" in role_label:
                role = 2
            else:
                continue
            code_map = mapping.setdefault(code, {})
            code_map[reviewer_name] = role
    return mapping


def parse_value_block(
    lines: List[str],
    start_index: int,
    field_names: Sequence[str] = FIELD_NAMES,
    field_aliases: Optional[Dict[str, str]] = None,
) -> tuple[str, str, int]:
    """Return (label, value, next_index) for the block starting at start_index."""
    raw_line = lines[start_index]
    label, value = raw_line.split(":", 1)
    label = label.strip()
    paragraphs: List[str] = []
    current: List[str] = []
    first_value = value.strip()
    if first_value:
        current.append(first_value)
    index = start_index + 1
    while index < len(lines):
        candidate = lines[index]
        stripped = candidate.strip()
        if ":" in stripped:
            if field_aliases:
                candidate_label = stripped.split(":", 1)[0].strip().lower()
                if candidate_label in field_aliases:
                    break
            else:
                if any(stripped.startswith(f"{name}:") for name in field_names):
                    break
        if stripped == "":
            if current:
                paragraphs.append(" ".join(current).strip())
                current = []
            index += 1
            continue
        if stripped.startswith("#"):
            if current:
                paragraphs.append(" ".join(current).strip())
                current = []
            internal_text = stripped[1:].strip()
            if paragraphs and paragraphs[-1].startswith("# "):
                paragraphs[-1] = paragraphs[-1] + " " + internal_text
            else:
                paragraphs.append("# " + internal_text)
        else:
            current.append(stripped)
        index += 1
    if current:
        paragraphs.append(" ".join(current).strip())
    cleaned_value = "\n\n".join(paragraphs).strip()
    return label, cleaned_value, index


def parse_simple_review_content(content: str, path: Path) -> Dict[str, ProposalSummary]:
    """Parse compact review files that list proposal codes and short fields."""
    lines = [line.rstrip() for line in content.splitlines()]
    code_indices = [
        idx for idx, line in enumerate(lines) if PROPOSAL_CODE_RE.match(line.strip())
    ]
    if not code_indices:
        return {}

    summaries: Dict[str, ProposalSummary] = {}
    reviewer_name = normalise_reviewer_from_filename(path)

    for pos, start_index in enumerate(code_indices):
        end_index = code_indices[pos + 1] if pos + 1 < len(code_indices) else len(lines)
        code = lines[start_index].strip()
        block_lines = lines[start_index + 1 : end_index]

        field_values = {name: "" for name in SIMPLE_FIELD_NAMES}
        index = 0
        while index < len(block_lines):
            current = block_lines[index]
            stripped = current.strip()
            if not stripped:
                index += 1
                continue
            if ":" not in stripped:
                index += 1
                continue

            label = stripped.split(":", 1)[0].strip()
            canonical = SIMPLE_FIELD_ALIASES.get(label.lower())
            if not canonical:
                index += 1
                continue

            label, value, index = parse_value_block(
                block_lines,
                index,
                field_names=SIMPLE_FIELD_NAMES,
                field_aliases=SIMPLE_FIELD_ALIASES,
            )
            field_values[canonical] = value

        if not any(field_values[name] for name in SIMPLE_FIELD_NAMES):
            continue

        summary = summaries.setdefault(
            code,
            ProposalSummary(
                code=code,
                title="",
                pi="",
                networks="",
                wavelengths="",
                reviews=[],
            ),
        )

        referee_comments = field_values["Referee comments"].strip()
        extra_sections: List[str] = []
        for label in ("General remark", "Strengths", "Weaknesses"):
            value = field_values.get(label, "").strip()
            if value:
                extra_sections.append(f"{label}: {value}")
        if extra_sections:
            if referee_comments:
                referee_comments = "\n\n".join([referee_comments] + extra_sections)
            else:
                referee_comments = "\n\n".join(extra_sections)

        review_entry = ReviewEntry(
            reviewer=reviewer_name,
            source_file=path,
            grade=field_values["Grade"],
            referee_comments=referee_comments,
            technical_review=field_values["Technical review"],
            time_recommended=field_values["Time recommended"],
        )
        summary.reviews.append(review_entry)

    return summaries


def parse_header_columns(header: str) -> tuple[str, str, str, str]:
    """Parse proposal header line into (code, pi, networks, wavelengths)."""
    stripped = header.strip()
    if stripped:
        columns = re.split(r"\s{2,}", stripped)
        if len(columns) >= 4 and PROPOSAL_CODE_RE.match(columns[0].strip()):
            code = columns[0].strip()
            pi = columns[1].strip()
            networks = " ".join(part.strip() for part in columns[2:-1] if part.strip())
            wavelengths = columns[-1].strip()
            return code, pi, networks, wavelengths

    # Fallback for legacy fixed-width templates.
    code = header[0:22].strip()
    pi = header[22:45].strip()
    networks = header[45:72].strip()
    wavelengths = header[72:].strip()
    return code, pi, networks, wavelengths


def parse_review_file(path: Path) -> Dict[str, ProposalSummary]:
    """Parse a single review file into proposal summaries keyed by proposal code."""
    content = read_text_with_fallback(path)
    if SEPARATOR not in content:
        simple_summaries = parse_simple_review_content(content, path)
        if simple_summaries:
            return simple_summaries

    blocks = [block.strip("\n") for block in content.split(SEPARATOR) if block.strip()]
    summaries: Dict[str, ProposalSummary] = {}

    for block in blocks:
        lines = [line.rstrip() for line in block.splitlines()]
        if not lines:
            continue

        header = lines[0]
        if not header.strip():
            continue

        exp, pi, networks, wavelengths = parse_header_columns(header)
        title = lines[1].strip() if len(lines) > 1 else ""

        field_values = {name: "" for name in FIELD_NAMES}
        index = 2
        while index < len(lines):
            current = lines[index]
            stripped = current.strip()
            if not stripped:
                index += 1
                continue
            if ":" not in stripped:
                index += 1
                continue

            label = stripped.split(":", 1)[0].strip()
            if label not in FIELD_NAMES:
                index += 1
                continue

            label, value, index = parse_value_block(lines, index)
            field_values[label] = value

        if not any(field_values[name] for name in FIELD_NAMES):
            continue  # Skip blocks without substantive content.

        if exp not in summaries:
            summaries[exp] = ProposalSummary(
                code=exp,
                title=title,
                pi=pi,
                networks=networks,
                wavelengths=wavelengths,
                reviews=[],
            )

        reviewer_name = normalise_reviewer_from_filename(path)
        review_entry = ReviewEntry(
            reviewer=reviewer_name,
            source_file=path,
            grade=field_values["Grade"],
            referee_comments=field_values["Referee comments"],
            technical_review=field_values["Technical review"],
            time_recommended=field_values["Time recommended"],
        )
        summaries[exp].reviews.append(review_entry)

    return summaries


def merge_summaries(files: Iterable[Path]) -> Dict[str, ProposalSummary]:
    """Merge per-file summaries into a combined dictionary keyed by proposal code."""
    combined: Dict[str, ProposalSummary] = {}
    for file_path in files:
        file_summaries = parse_review_file(file_path)
        for code, summary in file_summaries.items():
            if code not in combined:
                combined[code] = summary
            else:
                dest = combined[code]
                if not dest.title and summary.title:
                    dest.title = summary.title
                if not dest.pi and summary.pi:
                    dest.pi = summary.pi
                if not dest.networks and summary.networks:
                    dest.networks = summary.networks
                if not dest.wavelengths and summary.wavelengths:
                    dest.wavelengths = summary.wavelengths
                dest.reviews.extend(summary.reviews)
    return combined


def apply_assignments(
    summaries: Dict[str, ProposalSummary], assignments: Dict[str, Dict[str, int]]
) -> None:
    """Annotate review entries with assignment roles where available."""
    for code, summary in summaries.items():
        reviewers_map = assignments.get(code)
        if not reviewers_map:
            continue
        for review in summary.reviews:
            norm = normalise_person_name(review.reviewer)
            role = reviewers_map.get(norm)
            if role is not None:
                review.role = role


def load_conflicts(path: Path) -> Dict[str, str]:
    """Parse optional 'Conflicts:' section from the assignments summary file."""
    conflicts: Dict[str, str] = {}
    data = read_text_with_fallback(path)
    if not data.strip():
        return conflicts

    lines = data.splitlines()
    in_conflict_section = False
    conflict_pattern = re.compile(r"^\s*([A-Za-z0-9/_.-]+)\s*:\s*(.*?)\s*$")

    for raw_line in lines:
        line = raw_line.strip()
        if not in_conflict_section:
            if line.lower() == "conflicts:":
                in_conflict_section = True
            continue

        if not line:
            continue

        match = conflict_pattern.match(line)
        if not match:
            continue

        code = match.group(1).strip()
        value = match.group(2).strip()
        if not value or value.lower() == "none":
            conflicts[code] = "NONE"
        else:
            conflicts[code] = value

    return conflicts


def load_code_mapping(path: Path) -> Dict[str, str]:
    """Parse an EVN code mapping file into a dict {evn_code -> proposal_code}.

    Each non-blank line whose first whitespace-delimited token matches the
    proposal-code pattern (e.g. ``E26A001``) is treated as a mapping entry;
    the second token on that line is the alternative proposal code (e.g. ``EC107``).
    Lines that are continuations (PI name overflow etc.) are silently skipped.
    """
    mapping: Dict[str, str] = {}
    data = read_text_with_fallback(path)
    for line in data.splitlines():
        tokens = line.split()
        if len(tokens) >= 2 and PROPOSAL_CODE_RE.match(tokens[0]):
            mapping[tokens[0].upper()] = tokens[1]
    return mapping


def proposal_pre_grade(summary: ProposalSummary) -> str:
    """Return numeric mean grade across available reviewer grades."""
    numeric_values: List[float] = []
    for review in summary.reviews:
        numeric = parse_numeric_grade(review.grade.strip())
        if numeric is not None:
            numeric_values.append(numeric)
    if not numeric_values:
        return "N/A"
    return f"{(sum(numeric_values) / len(numeric_values)):.2f}"


def proposal_grade_std(summary: ProposalSummary) -> str:
    """Return population std-dev across available reviewer grades."""
    numeric_values: List[float] = []
    for review in summary.reviews:
        numeric = parse_numeric_grade(review.grade.strip())
        if numeric is not None:
            numeric_values.append(numeric)
    if not numeric_values:
        return "N/A"

    mean = sum(numeric_values) / len(numeric_values)
    variance = sum((value - mean) ** 2 for value in numeric_values) / len(
        numeric_values
    )
    return f"{math.sqrt(variance):.2f}"


def proposal_grade_std_numeric(summary: ProposalSummary) -> Optional[float]:
    """Return numeric population std-dev across available reviewer grades."""
    numeric_values: List[float] = []
    for review in summary.reviews:
        numeric = parse_numeric_grade(review.grade.strip())
        if numeric is not None:
            numeric_values.append(numeric)
    if not numeric_values:
        return None

    mean = sum(numeric_values) / len(numeric_values)
    variance = sum((value - mean) ** 2 for value in numeric_values) / len(
        numeric_values
    )
    return math.sqrt(variance)


def proposal_low_grade_count(summary: ProposalSummary, threshold: float = 2.0) -> int:
    """Count numeric reviewer grades greater than or equal to threshold."""
    count = 0
    for review in summary.reviews:
        numeric = parse_numeric_grade(review.grade.strip())
        if numeric is not None and numeric >= threshold:
            count += 1
    return count


def proposal_role_initials(
    summary: ProposalSummary, assignments: Dict[str, Dict[str, int]]
) -> str:
    """Return '<primary>/<secondary>' initials for agenda output."""
    primary: Optional[str] = None
    secondary: Optional[str] = None

    for review in summary.reviews:
        if review.role == 1 and primary is None:
            primary = reviewer_initials(review.reviewer)
        elif review.role == 2 and secondary is None:
            secondary = reviewer_initials(review.reviewer)

    code_map = assignments.get(summary.code) or assignments.get(summary.code.upper()) or {}
    if primary is None or secondary is None:
        for reviewer_name, role in code_map.items():
            if role == 1 and primary is None:
                primary = reviewer_initials(reviewer_name)
            elif role == 2 and secondary is None:
                secondary = reviewer_initials(reviewer_name)

    return f"{primary or '?'}/{secondary or '?'}"


def build_agenda_text(
    summaries: Dict[str, ProposalSummary],
    assignments: Dict[str, Dict[str, int]],
    conflicts: Dict[str, str],
) -> str:
    """Render agenda-friendly text entries (one two-line block per proposal)."""
    blocks: List[str] = []
    std_by_code = {
        code: proposal_grade_std_numeric(summary) for code, summary in summaries.items()
    }

    # Sort by grade std-dev (descending), then by proposal code for stable ties.
    sorted_codes = sorted(
        summaries.keys(),
        key=lambda code: (
            -(std_by_code[code] if std_by_code[code] is not None else -1.0),
            code,
        ),
    )

    for code in sorted_codes:
        summary = summaries[code]
        header = summary.code
        if summary.pi:
            header += f" {summary.pi}"
        if summary.title:
            header += f" - {summary.title}"

        role_initials = proposal_role_initials(summary, assignments)
        pre_grade = proposal_pre_grade(summary)
        std_grade = proposal_grade_std(summary)
        low_grade_count = proposal_low_grade_count(summary, threshold=2.0)
        conflict_value = (
            conflicts.get(summary.code)
            or conflicts.get(summary.code.upper())
            or conflicts.get(code)
            or conflicts.get(code.upper())
            or "NONE"
        )

        blocks.append(header.strip())
        blocks.append(
            f"[{role_initials}; pre-grade: {pre_grade}, std: {std_grade}, N>2.0: {low_grade_count}] CONFLICT: {conflict_value}"
        )
        blocks.append("")

    return "\n".join(blocks).rstrip() + "\n"


def process_text_for_latex(text: str) -> str:
    """Escape and render review text, highlighting #-prefixed lines as internal PC comments.

    Lines whose first non-whitespace character is ``#`` are treated as
    internal committee notes.  They are stripped of the leading ``#``,
    prefixed with **Internal:** and rendered in blue.  All other lines are
    normalised and escaped as usual.
    """
    if not text:
        return text

    lines = text.splitlines()
    segments: List[tuple[str, str]] = []
    current_regular: List[str] = []

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("#"):
            if current_regular:
                segments.append(("regular", "\n".join(current_regular)))
                current_regular = []
            segments.append(("internal", stripped[1:].strip()))
        else:
            current_regular.append(line)
    if current_regular:
        segments.append(("regular", "\n".join(current_regular)))

    result_parts: List[str] = []
    for kind, content in segments:
        if kind == "regular":
            normalized = normalize_paragraphs(content)
            if normalized:
                result_parts.append(latex_escape(normalized))
        else:
            escaped = latex_escape(content)
            result_parts.append(
                r"{\color{blue}\textbf{Internal:} " + escaped + "}"
            )

    return "\n\n".join(part for part in result_parts if part)


def format_review_block(review: ReviewEntry) -> str:
    """Return LaTeX formatted block for a single review entry."""
    parts: List[str] = []
    reviewer_label = latex_escape(review.reviewer)
    supers = role_superscript(review.role)
    parts.append(f"\\subsection*{{Reviewer: {reviewer_label}{supers}}}")

    if review.grade:
        parts.append(f"\\textbf{{Grade:}} {latex_escape(review.grade)}\\\\")
    if review.time_recommended:
        parts.append(
            f"\\textbf{{Time recommended:}} {latex_escape(review.time_recommended)}\\\\"
        )

    if review.referee_comments:
        parts.append("\\textbf{Referee comments}\\\\")
        comments = process_text_for_latex(review.referee_comments)
        parts.append(comments)

    if review.technical_review:
        parts.append("\\par")
        parts.append("\\textbf{Technical review}\\\\")
        tech = process_text_for_latex(review.technical_review)
        parts.append(tech)

    source_path = latex_escape(str(review.source_file))
    parts.append("\\par")
    parts.append(f"\\textit{{Source file:}} {source_path}\\\\")
    return "\n".join(parts)


def build_latex_document(
    summaries: Dict[str, ProposalSummary],
    title: str,
    version: str = "",
    code_mapping: Optional[Dict[str, str]] = None,
) -> str:
    """Render the combined summaries into a LaTeX document string."""
    preamble = [
        r"\documentclass{article}",
        r"\usepackage[T1]{fontenc}",
        r"\usepackage[utf8]{inputenc}",
        r"\usepackage{lmodern}",
        r"\usepackage[margin=2.5cm]{geometry}",
        r"\usepackage[hidelinks]{hyperref}",
        r"\usepackage{longtable}",
        r"\usepackage{enumitem}",
        r"\usepackage{xcolor}",
        r"\usepackage{multicol}",
        r"\usepackage{titling}",
        r"\setlength{\droptitle}{-1.5em}",
        r"\predate{\vspace{-1em}\begin{center}}",
        r"\postdate{\par\end{center}}",
        r"\renewcommand{\familydefault}{\sfdefault}",
        r"\setlength{\parindent}{0pt}",
        r"\setlength{\parskip}{5pt}",
        r"\setlist{nosep}",
        *[
            rf"\DeclareUnicodeCharacter{{{codepoint}}}{{{latex}}}"
            for codepoint, latex in UNICODE_LATEX_DECLARATIONS.items()
        ],
        r"\begin{document}",
        rf"\title{{{latex_escape(title)}}}",
    ]

    date_parts: List[str] = []
    if version:
        date_parts.append(f"Version~{latex_escape(version)}")
    date_parts.append(r"\today")
    date_content = r" \\ ".join(date_parts)

    preamble.extend(
        [
            rf"\date{{{date_content}}}",
            r"\maketitle",
            r"\tableofcontents",
            r"\clearpage",
        ]
    )
    body: List[str] = []

    for idx, code in enumerate(sorted(summaries.keys())):
        summary = summaries[code]
        grade_entries = [
            (reviewer_initials(review.reviewer), review.grade.strip(), review.role)
            for review in summary.reviews
            if review.grade.strip()
        ]
        if grade_entries:
            sorted_entries = sorted(
                grade_entries, key=lambda item: (role_sort_key(item[2]), item[0])
            )
            header_cells: List[str] = []
            values: List[str] = []
            numeric_values: List[float] = []
            for initials, grade, role in sorted_entries:
                header_cells.append(f"{latex_escape(initials)}{role_superscript(role)}")
                display_grade = "4.0" if SPECIAL_GRADE_RE.search(grade or "") else grade
                values.append(display_grade)
                numeric = parse_numeric_grade(grade)
                if numeric is not None:
                    numeric_values.append(numeric)
            if numeric_values:
                mean = sum(numeric_values) / len(numeric_values)
                variance = sum((value - mean) ** 2 for value in numeric_values) / len(
                    numeric_values
                )
                average_value = f"{mean:.2f}"
                std_dev_value = f"{math.sqrt(variance):.2f}"
            else:
                average_value = "N/A"
                std_dev_value = "N/A"
            header_cells.append(latex_escape("Average"))
            header_cells.append(latex_escape("Std Dev"))
            values.append(average_value)
            values.append(std_dev_value)

            header_row = " & ".join(header_cells)
            value_row_parts = []
            summary_start_index = len(sorted_entries)
            for idx_col, value in enumerate(values):
                escaped = latex_escape(value)
                if idx_col >= summary_start_index:
                    escaped = f"\\textbf{{{escaped}}}"
                value_row_parts.append(escaped)
            value_row = " & ".join(value_row_parts)

            table_lines = [
                "\\begin{center}",
                f"\\begin{{tabular}}{{{'c' * len(header_cells)}}}",
                header_row + r"\\",
                r"\hline",
                value_row + r"\\",
                "\\end{tabular}",
                "\\end{center}",
            ]
        else:
            table_lines = []

        if idx > 0:
            body.append("\\clearpage")

        alt_code = (code_mapping or {}).get(summary.code.upper(), "")
        code_label = f"{summary.code}/{alt_code}" if alt_code else summary.code
        section_title = f"{code_label}: {summary.title}" if summary.title else code_label
        body.append(f"\\section*{{{latex_escape(section_title)}}}")
        body.append(f"\\addcontentsline{{toc}}{{section}}{{{latex_escape(section_title)}}}")
        if summary.pi:
            body.append(f"\\textbf{{PI:}} {latex_escape(summary.pi)}\\\\")
        if summary.networks:
            body.append(f"\\textbf{{Networks:}} {latex_escape(summary.networks)}\\\\")
        if summary.wavelengths:
            body.append(
                f"\\textbf{{Requested wavelengths:}} {latex_escape(summary.wavelengths)}\\\\"
            )
        if table_lines:
            body.append("")
            body.extend(table_lines)
        if table_lines:
            body.append("")
        body.append("")

        if not summary.reviews:
            body.append("\\textit{No reviews available.}")
            continue

        body.append("\\vspace{10pt}")
        body.append("\\begin{multicols}{2}")
        # Sort reviews: primary, secondary, then the rest (by name).
        for review in sorted(
            summary.reviews,
            key=lambda r: (role_sort_key(r.role), r.reviewer.lower()),
        ):
            body.append(format_review_block(review))
            body.append("")  # Blank line between reviews
        body.append("\\end{multicols}")

    footer = [r"\end{document}"]
    parts: List[str] = []
    parts.extend(preamble)
    parts.extend(body)
    parts.extend(footer)
    return "\n".join(parts)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)

    reviews_dir = args.reviews_dir
    if not reviews_dir.exists() or not reviews_dir.is_dir():
        print(f"Reviews directory not found: {reviews_dir}", file=sys.stderr)
        return 1

    review_files = sorted(reviews_dir.glob("*.txt"))
    if not review_files:
        print(f"No .txt review files found in {reviews_dir}", file=sys.stderr)
        return 1

    combined = merge_summaries(review_files)
    if not combined:
        print("No proposal data extracted from review files.", file=sys.stderr)
        return 1

    assignment_map: Dict[str, Dict[str, int]] = {}
    conflict_map: Dict[str, str] = {}
    if args.assignments:
        if not args.assignments.is_file():
            print(f"Assignments file not found: {args.assignments}", file=sys.stderr)
            return 1
        assignment_map = load_assignments(args.assignments)
        conflict_map = load_conflicts(args.assignments)
    if assignment_map:
        apply_assignments(combined, assignment_map)

    code_mapping: Dict[str, str] = {}
    if args.code_mapping:
        if not args.code_mapping.is_file():
            print(f"Code mapping file not found: {args.code_mapping}", file=sys.stderr)
            return 1
        code_mapping = load_code_mapping(args.code_mapping)

    latex_content = build_latex_document(combined, args.title, args.version, code_mapping)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(latex_content, encoding="utf-8")
    print(f"Wrote LaTeX summary to {args.output}")

    if args.agenda_txt:
        agenda_content = build_agenda_text(combined, assignment_map, conflict_map)
        args.agenda_txt.parent.mkdir(parents=True, exist_ok=True)
        args.agenda_txt.write_text(agenda_content, encoding="utf-8")
        print(f"Wrote agenda text summary to {args.agenda_txt}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
