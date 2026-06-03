#!/usr/bin/env python3
"""Parse EVN proposal PDFs and render chair summaries.

Usage
=====
python proposal_to_review_template.py [-h] [-p PDF_DIR] [-o OUTPUT] [-m PC_MEMBERS] [-a ASSIGNMENTS] [pdfs ...]

positional arguments:
  pdfs                  Specific PDF files to process. Defaults to all PDFs in --pdf-dir.

optional arguments:
  -h, --help            Show this help message and exit.
  -p PDF_DIR, --pdf-dir PDF_DIR
                        Directory containing proposal PDFs (required if no PDF files are listed).
  -o OUTPUT, --output OUTPUT
                        Write the formatted output to this file instead of stdout.
  -m PC_MEMBERS, --pc-members PC_MEMBERS
                        File containing EVN PC members (one per line) to auto-assign reviewers. If a member
                        should always referee a specific proposal, append tokens like CODE#1 (primary reviewer)
                        or CODE#2 (secondary reviewer) after their name, e.g. "Jane Smith E25A001#1 E25A010#2".
  -a ASSIGNMENTS, --assignments ASSIGNMENTS
                        File to store reviewer assignments (defaults to reviewer_assignments.txt when --pc-members is used).
  --reviewers-per-proposal COUNT
                        Total reviewers to assign per proposal (minimum 2, default 2).
  --max-per-member MAX  Maximum number of proposals assigned to any single PC member.
  --max-first-per-member COUNT
                        Maximum number of primary-reviewer slots per PC member.
  --max-second-per-member COUNT
                        Maximum number of secondary-reviewer slots per PC member.
  --member-summary FILE Write a per-member HTML assignment table to FILE.
  --conflicts-file FILE Load additional conflicts from FILE (same format as reviewer_assignments appendix).
  --science-tags-file FILE
                        Write inferred science categories per proposal to FILE.
  --prefer-matching-tags
                        Prefer assigning primary reviewers whose expertise tags match the proposal's science tags.

Examples:
  python proposal_to_review_template.py -p test_proposals -m EVN_pc_members.txt
  python proposal_to_review_template.py proposal.pdf another.pdf -m EVN_pc_members.txt -o summaries.txt
"""

from __future__ import annotations

import argparse
import html
import re
import subprocess
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple, Union
from xml.sax.saxutils import escape
from collections import defaultdict

from template import render_record

DocxRecord = Dict[str, Any]
VERBOSE = True

# Match experiment identifiers and waveband tokens in the PDF text.
PROPOSAL_CODE_RE = re.compile(r"\b[EG]\d{2}[A-Z]\d{3}\b")
WAVEBAND_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(cm|mm|m|GHz|MHz)", re.IGNORECASE)


def log_verbose(message: str) -> None:
    if VERBOSE:
        print(f"[INFO] {message}", file=sys.stderr)


def generate_role_labels(count: int) -> List[str]:
    """Return role labels for the ordered reviewer slots."""
    labels: List[str] = []
    for idx in range(count):
        if idx == 0:
            labels.append("Primary Reviewer")
        elif idx == 1:
            labels.append("Secondary Reviewer")
        else:
            labels.append(f"Additional Review {idx - 1}")
    return labels

# Known network keywords to capture from the summary table.
NETWORK_KEYWORDS = {
    "EVN",
    "MERLIN",
    "VLBA",
    "MeerKAT",
    "uGMRT",
    "NRAO",
    "LBA",
    "EAVN",
    "KVN",
    "GMVA",
    "ATCA",
    "Other",
    "JVN",
    "JNET",
    "e-MERLIN",
}
SUMMARY_STOP_PREFIXES = (
    "no phd",
    "students involved",
    "student",
    "is this",
    "linked proposal",
    "relevant previous",
    "observation dependencies",
    "aggregate correlator",
    "processor information",
    "print view prepared",
    "scientific category",
    "scheduling assistance",
    "rapid response science",
)
SUMMARY_SKIP_PHRASES = (
    "Observation Number of Network",
    "number      targets",
    "Aggregate Correlator e-EVN",
    "Out-of-",
)

# Lowercase salutations and honorifics to strip from normalised names.
TITLE_PREFIXES = {
    "dr",
    "prof",
    "professor",
    "mr",
    "mrs",
    "ms",
    "miss",
    "sir",
    "madam",
}

SCIENCE_CATEGORY_RULES: Dict[str, Tuple[Tuple[str, str, int], ...]] = {
    "Galactic": (
        ("keyword", "galactic", 2),
        ("keyword", "milky way", 3),
        ("keyword", "stellar", 1),
        ("keyword", "star formation", 2),
        ("keyword", "protostar", 2),
        ("keyword", "brown dwarf", 2),
        ("keyword", "molecular cloud", 2),
        ("regex", r"\bpulsar(s)?\b", 3),
        ("regex", r"\bpsr\b", 3),
        ("regex", r"\b(binary|compact) star(s)?\b", 2),
    ),
    "Extragalactic": (
        ("keyword", "extragalactic", 3),
        ("keyword", "galaxy", 1),
        ("keyword", "galaxies", 1),
        ("keyword", "merger", 1),
        ("keyword", "cluster", 1),
        ("keyword", "intergalactic", 2),
        ("keyword", "lensing", 2),
        ("regex", r"\bjet(s)?\b", 1),
        ("regex", r"\bradio galaxy\b", 3),
    ),
    "Spectral Line": (
        ("keyword", "spectral line", 3),
        ("keyword", "line emission", 2),
        ("keyword", "absorption line", 2),
        ("keyword", "maser", 3),
        ("keyword", "masers", 3),
        ("keyword", "molecular line", 2),
        ("keyword", "molecular transition", 2),
        ("keyword", "co(1-0)", 3),
        ("keyword", "co(2-1)", 3),
        ("keyword", "ammonia", 2),
        ("keyword", "hi ", 2),
        ("keyword", "h i", 2),
        ("regex", r"\b21\s*cm\b", 2),
    ),
    "Maser": (
        ("keyword", "maser", 4),
        ("keyword", "masers", 4),
        ("keyword", "megamaser", 4),
        ("keyword", "ohm", 3),
        ("keyword", "oh maser", 4),
        ("keyword", "water maser", 4),
        ("keyword", "h2o maser", 4),
        ("keyword", "methanol maser", 4),
        ("keyword", "lirg", 2),
        ("keyword", "ulirg", 2),
    ),
    "Transient": (
        ("keyword", "transient", 3),
        ("keyword", "transients", 3),
        ("keyword", "burst", 2),
        ("keyword", "afterglow", 2),
        ("keyword", "flare", 1),
        ("keyword", "tidal disruption", 3),
        ("keyword", "fast radio burst", 4),
        ("keyword", "frb", 3),
        ("keyword", "grb", 3),
    ),
    "AGN": (
        ("keyword", "agn", 3),
        ("keyword", "active galactic nucleus", 4),
        ("keyword", "active galactic nuclei", 4),
        ("keyword", "blazar", 3),
        ("keyword", "seyfert", 3),
        ("keyword", "quasar", 3),
        ("keyword", "core-jet", 2),
        ("regex", r"\bagns\b", 2),
    ),
    "Supernovae": (
        ("keyword", "supernova", 4),
        ("keyword", "supernovae", 4),
        ("regex", r"\bsn\s?\d+", 3),
        ("keyword", "snr", 3),
        ("keyword", "remnant", 2),
        ("keyword", "nova remnant", 3),
    ),
    "Pulsar": (
        ("regex", r"\bpulsar(s)?\b", 4),
        ("regex", r"\bpsr\b", 3),
        ("keyword", "pulsar timing", 4),
        ("keyword", "scintillation", 3),
        ("keyword", "dispersion measure", 3),
        ("keyword", "magnetar", 3),
        ("keyword", "neutron star", 2),
        ("keyword", "millisecond pulsar", 3),
        ("keyword", "msp", 2),
        ("keyword", "timing noise", 2),
        ("keyword", "interstellar medium", 1),
    ),
    "Astrometry": (
        ("keyword", "astrometry", 4),
        ("keyword", "proper motion", 3),
        ("keyword", "parallax", 3),
        ("keyword", "reference frame", 3),
        ("keyword", "icrf", 3),
        ("keyword", "geodesy", 3),
        ("keyword", "tropospheric", 2),
        ("regex", r"\b(μas|microarcsecond)\b", 3),
        ("regex", r"\bmas\s+precision\b", 2),
    ),
}
SCIENCE_CATEGORY_ALIAS_TERMS: Dict[str, str] = {
    "galactic": "Galactic",
    "milky way": "Galactic",
    "stellar": "Galactic",
    "star formation": "Galactic",
    "starformation": "Galactic",
    "pulsar": "Galactic",
    "psr": "Galactic",
    "extragalactic": "Extragalactic",
    "galaxy": "Extragalactic",
    "galaxies": "Extragalactic",
    "cluster": "Extragalactic",
    "spectral line": "Spectral Line",
    "spectralline": "Spectral Line",
    "line emission": "Spectral Line",
    "molecular line": "Spectral Line",
    "molecularline": "Spectral Line",
    "hi": "Spectral Line",
    "maser": "Maser",
    "masers": "Maser",
    "megamaser": "Maser",
    "ohm": "Maser",
    "oh maser": "Maser",
    "water maser": "Maser",
    "methanol maser": "Maser",
    "transient": "Transient",
    "transients": "Transient",
    "burst": "Transient",
    "frb": "Transient",
    "grb": "Transient",
    "agn": "AGN",
    "agns": "AGN",
    "active galactic nucleus": "AGN",
    "blazar": "AGN",
    "seyfert": "AGN",
    "quasar": "AGN",
    "supernova": "Supernovae",
    "supernovae": "Supernovae",
    "snr": "Supernovae",
    "remnant": "Supernovae",
    "pulsar": "Pulsar",
    "pulsars": "Pulsar",
    "psr": "Pulsar",
    "scintillation": "Pulsar",
    "magnetar": "Pulsar",
    "neutron star": "Pulsar",
    "millisecond pulsar": "Pulsar",
    "msp": "Pulsar",
    "astrometry": "Astrometry",
    "proper motion": "Astrometry",
    "parallax": "Astrometry",
    "reference frame": "Astrometry",
    "icrf": "Astrometry",
    "geodesy": "Astrometry",
    "other": "Other",
}
SCIENCE_CATEGORY_MIN_SCORE = 3
SCIENCE_DEFAULT_CATEGORY = "Other"


def normalise_name(name: str) -> str:
    """Return a lowercase, punctuation-free version of a personal name."""
    cleaned = re.sub(r"\s+", " ", name).strip().lower()
    if not cleaned:
        return ""
    cleaned = re.sub(r"[^\w\s]", "", cleaned)
    tokens = cleaned.split()
    filtered_tokens = [token for token in tokens if token not in TITLE_PREFIXES]
    return " ".join(filtered_tokens) if filtered_tokens else cleaned


def infer_science_categories(title: str, lines: Sequence[str]) -> List[str]:
    """Return a list of science categories inferred from free text and declared fields."""
    declared = _extract_declared_science_categories(lines)

    haystack_parts = [title.lower()]
    haystack_parts.extend(line.lower() for line in lines if line.strip())
    haystack = " ".join(haystack_parts)
    scores: Dict[str, int] = {category: 0 for category in SCIENCE_CATEGORY_RULES}

    for category, rules in SCIENCE_CATEGORY_RULES.items():
        for rule_type, pattern, weight in rules:
            if rule_type == "keyword":
                if pattern in haystack:
                    scores[category] += weight
            elif rule_type == "regex":
                if re.search(pattern, haystack):
                    scores[category] += weight

    if declared:
        # Also add any category whose keyword fires strongly in the title alone,
        # so a title like "Pulsar timing..." isn't lost when the declared field says "Galactic".
        title_lower = title.lower()
        title_scores: Dict[str, int] = {category: 0 for category in SCIENCE_CATEGORY_RULES}
        for category, rules in SCIENCE_CATEGORY_RULES.items():
            for rule_type, pattern, weight in rules:
                if rule_type == "keyword":
                    if pattern in title_lower:
                        title_scores[category] += weight
                elif rule_type == "regex":
                    if re.search(pattern, title_lower):
                        title_scores[category] += weight
        title_extras = [cat for cat, value in title_scores.items() if value >= SCIENCE_CATEGORY_MIN_SCORE]
        return sorted(set(declared) | set(title_extras))

    selected = [cat for cat, value in scores.items() if value >= SCIENCE_CATEGORY_MIN_SCORE]
    if not selected:
        max_score = max(scores.values(), default=0)
        if max_score > 0:
            selected = [cat for cat, value in scores.items() if value == max_score]
    if not selected:
        selected = [SCIENCE_DEFAULT_CATEGORY]
    return sorted(set(selected))


def _extract_declared_science_categories(lines: Sequence[str]) -> List[str]:
    categories: List[str] = []
    pending = False
    for raw_line in lines:
        line = raw_line.strip()
        lower = line.lower()
        if pending:
            categories.extend(_science_aliases_from_text(lower))
            pending = False
            continue
        if "scientific category" in lower or "science category" in lower or lower.startswith("category"):
            if ":" in line:
                _, _, remainder = line.partition(":")
                categories.extend(_science_aliases_from_text(remainder.lower()))
            else:
                pending = True
    return sorted(set(categories))


def _science_aliases_from_text(text: str) -> List[str]:
    categories: List[str] = []
    parts = re.split(r"[;,/]| and |\s{2,}", text)
    for part in parts:
        chunk = part.strip()
        if not chunk:
            continue
        for alias, category in SCIENCE_CATEGORY_ALIAS_TERMS.items():
            if alias in chunk:
                categories.append(category)
                break
    return categories


class PdfExtractError(RuntimeError):
    """Raised when pdftotext is unavailable or fails."""


def extract_pdf_lines(path: Path) -> List[str]:
    """Return the PDF content as layout-preserving text lines via pdftotext."""
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", str(path), "-"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except FileNotFoundError as exc:  # pragma: no cover - poppler not installed
        raise PdfExtractError("pdftotext command not found; please install Poppler.") from exc
    except subprocess.CalledProcessError as exc:
        stderr = exc.stderr.decode("utf-8", errors="ignore") if exc.stderr else ""
        raise PdfExtractError(f"pdftotext failed for {path.name}: {stderr.strip()}") from exc
    text = result.stdout.decode("utf-8", errors="ignore")
    return text.splitlines()


def find_experiment_and_title(lines: Sequence[str], fallback: str) -> tuple[str, str, Optional[str]]:
    """Locate the proposal code, title, and inline PI hint from text lines."""
    exp = None
    exp_idx = 0
    pi_hint = None
    for idx, line in enumerate(lines):
        match = PROPOSAL_CODE_RE.search(line)
        if match:
            exp = match.group()
            exp_idx = idx
            pi_hint = line[: match.start()].strip() or None
            break
    if exp is None:
        return fallback, fallback, None

    title_parts: List[str] = []
    for line in lines[exp_idx + 1 :]:
        stripped = line.strip()
        if not stripped:
            if title_parts:
                break
            continue
        if stripped.lower().startswith("abstract"):
            break
        title_parts.append(stripped)
    title = " ".join(title_parts).strip() or fallback
    return exp, title, pi_hint


def parse_applicants(lines: Sequence[str]) -> List[dict[str, str]]:
    """Parse the Applicants table into a list of dictionaries."""
    try:
        start = lines.index("Applicants")
    except ValueError:
        return []

    rows: List[dict[str, str]] = []
    current: Optional[dict[str, str]] = None
    keys = ["name", "affiliation", "email", "country", "potential"]
    skip_tokens = {"observer", "potential"}

    i = start + 1
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            i += 1
            continue
        if stripped == "Contact Author":
            if current:
                rows.append(current)
            break
        if stripped.lower() in skip_tokens:
            i += 1
            continue
        if "Name" in stripped and "Affiliation" in stripped and "Potential" in stripped:
            i += 1
            continue

        parts = [p for p in re.split(r"\s{2,}", line.rstrip()) if p]
        if not parts:
            i += 1
            continue

        if len(parts) >= len(keys):
            if current:
                rows.append(current)
            current = {keys[idx]: parts[idx].strip() for idx in range(len(keys))}
        elif len(parts) == len(keys) - 1:
            # Row with one column absent (e.g. no Potential Observer entry).
            if current:
                rows.append(current)
            current = {key: "" for key in keys}
            for idx in range(len(parts)):
                current[keys[idx]] = parts[idx].strip()
        else:
            if current is None:
                current = {key: "" for key in keys}
            if len(parts) == 1:
                fragment = parts[0].strip()
                email_so_far = current.get("email", "")
                # A wrapped domain suffix (e.g. "uk" or "ter.ac.uk" from
                # "postgrad.manchester.ac.uk") should be appended to a partial email
                # rather than treated as an affiliation continuation.
                # Heuristic: all-lowercase with dots/hyphens/digits, and either
                # contains a "." or is short (≤4 chars) like a bare TLD.
                if (email_so_far and "@" in email_so_far
                        and re.match(r'^\.?[a-z0-9][a-z0-9.-]*$', fragment)
                        and ("." in fragment or len(fragment.lstrip(".")) <= 4)):
                    # Direct concatenation — PDF column truncates mid-character,
                    # so no separator is needed; the fragment picks up exactly
                    # where the truncated string left off.
                    current["email"] = email_so_far + fragment
                else:
                    current["affiliation"] = " ".join(filter(None, [current.get("affiliation"), fragment]))
            elif len(parts) == 2:
                current["affiliation"] = " ".join(filter(None, [current.get("affiliation"), parts[0].strip()]))
                email_so_far = current.get("email", "")
                part1 = parts[1].strip()
                # Same domain-suffix heuristic as the 1-part branch: if the
                # current email is partial and parts[1] looks like a domain
                # fragment (no @, all lowercase/dots/hyphens), concatenate.
                if (email_so_far and "@" in email_so_far
                        and re.match(r'^\.?[a-z0-9][a-z0-9.-]*$', part1)
                        and ("." in part1 or len(part1.lstrip(".")) <= 4)):
                    current["email"] = email_so_far + part1
                else:
                    current["email"] = " ".join(filter(None, [email_so_far, part1]))
            else:
                current["potential"] = " ".join(filter(None, [current.get("potential"), " ".join(parts).strip()]))
        i += 1

    if current and (not rows or current is not rows[-1]):
        rows.append(current)
    return rows


def parse_contact_name(lines: Sequence[str]) -> Optional[str]:
    """Extract the contact author name from the dedicated section."""
    try:
        start = lines.index("Contact Author")
    except ValueError:
        return None

    for idx in range(start + 1, min(start + 10, len(lines))):
        stripped = lines[idx].strip()
        if not stripped:
            continue
        parts = [p for p in re.split(r"\s{2,}", stripped) if p]
        if not parts:
            continue
        if parts[0].lower() == "name" and len(parts) > 1:
            return parts[1].strip()
    return None


def _best_email(raw: str) -> Optional[str]:
    """Return the most plausible email address from a (potentially garbled) string.

    PDF layout wrapping can concatenate multiple fragments into the email field,
    e.g. 'ter.ac.uk 7@manchester.ac.uk lasheras@soton.ac.uk'.  We pick the
    candidate whose local part (before '@') is longest, as short local parts
    like '7' are artefacts of a wrapped previous token.
    """
    candidates: List[Tuple[int, str]] = []
    for token in raw.split():
        token = token.strip(".,;()<>[]")
        if "@" not in token:
            continue
        local, _, domain = token.partition("@")
        if local and "." in domain:
            candidates.append((len(local), token))
    if not candidates:
        return None
    candidates.sort(key=lambda x: -x[0])
    return candidates[0][1]


def parse_contact_email(lines: Sequence[str]) -> Optional[str]:
    """Extract the contact author email from the dedicated Contact Author section."""
    try:
        start = lines.index("Contact Author")
    except ValueError:
        return None

    for idx in range(start + 1, min(start + 15, len(lines))):
        stripped = lines[idx].strip()
        if not stripped:
            continue
        parts = [p for p in re.split(r"\s{2,}", stripped) if p]
        if not parts:
            continue
        if parts[0].lower() == "email" and len(parts) > 1:
            return _best_email(parts[1].strip()) or None
    return None


def normalise_waveband(value: str, unit: str) -> str:
    """Standardise waveband labels with consistent units."""
    unit_map = {"cm": "cm", "mm": "mm", "m": "m", "ghz": "GHz", "mhz": "MHz"}
    normalised_unit = unit_map.get(unit.lower(), unit.upper())
    clean_value = value[:-2] if value.endswith(".0") else value
    return f"{clean_value} {normalised_unit}"


def extract_wavebands(segment: str, collected: List[str]) -> None:
    """Collect unique waveband entries from a summary segment."""
    for value, unit in WAVEBAND_RE.findall(segment):
        label = normalise_waveband(value, unit)
        if label not in collected:
            collected.append(label)


def segments_to_network_tokens(segments: Sequence[str]) -> List[str]:
    """Flatten multi-line network segments into distinct network tokens."""
    tokens: List[str] = []
    buffer = ""
    for segment in segments:
        segment = segment.strip()
        if not segment or segment in {",", ";"}:
            continue
        parts = [p.strip() for p in segment.split(",") if p.strip()]
        trailing_comma = segment.endswith(",")
        if not parts:
            if trailing_comma and buffer:
                if keep_network(buffer) and buffer not in tokens:
                    tokens.append(buffer)
                buffer = ""
            continue
        for idx, part in enumerate(parts):
            if buffer:
                if part.upper() == "NRAO" and buffer.lower().endswith("other"):
                    buffer = f"{buffer} {part}"
                elif part.islower() or part == part.lower():
                    buffer = f"{buffer} {part}".strip()
                else:
                    if keep_network(buffer) and buffer not in tokens:
                        tokens.append(buffer)
                    buffer = part
            else:
                buffer = part
            if idx < len(parts) - 1 or trailing_comma:
                if keep_network(buffer) and buffer not in tokens:
                    tokens.append(buffer)
                buffer = ""
    if buffer and keep_network(buffer) and buffer not in tokens:
        tokens.append(buffer)
    return tokens


def keep_network(token: str) -> bool:
    """Return True if the token contains a recognized network keyword."""
    upper = token.upper()
    return any(keyword in upper for keyword in NETWORK_KEYWORDS)


def strip_numbers(segment: str) -> str:
    """Remove numeric values and waveband text to leave network fragments."""
    cleaned = segment
    for value, unit in WAVEBAND_RE.findall(segment):
        cleaned = cleaned.replace(f"{value}{unit}", "")
        cleaned = cleaned.replace(f"{value} {unit}", "")
    cleaned = re.sub(r"\d+(?:\.\d+)?", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip()


def parse_summary(lines: Sequence[str]) -> tuple[List[str], List[str]]:
    """Pull network list and requested wavebands from the summary table."""
    try:
        start = lines.index("Summary of Observations")
    except ValueError:
        return [], []

    networks: List[str] = []
    wavebands: List[str] = []
    current_segments: List[str] = []

    for line in lines[start + 1 :]:
        stripped = line.strip()
        if not stripped:
            continue
        lower = stripped.lower()
        if any(lower.startswith(prefix) for prefix in SUMMARY_STOP_PREFIXES):
            break
        if any(phrase in stripped for phrase in SUMMARY_SKIP_PHRASES):
            continue

        parts = [p for p in re.split(r"\s{2,}", stripped) if p]
        new_row = parts and parts[0].isdigit()
        data_parts: Iterable[str] = parts[2:] if new_row else parts

        if new_row and current_segments:
            for token in segments_to_network_tokens(current_segments):
                if token not in networks:
                    networks.append(token)
            current_segments = []

        for idx, segment in enumerate(data_parts):
            if not segment.strip():
                continue
            extract_wavebands(segment, wavebands)
            if new_row and idx > 0:
                continue
            cleaned = strip_numbers(segment)
            if cleaned:
                current_segments.append(cleaned)

    if current_segments:
        for token in segments_to_network_tokens(current_segments):
            if token not in networks:
                networks.append(token)

    return networks, wavebands


def load_pc_members(path: Path) -> Tuple[List[str], Dict[str, Dict[str, List[str]]], Dict[str, str], Set[str], Dict[str, Set[str]], Dict[str, Dict[str, int]]]:
    """Read PC member entries and return names, fixed reviewer preferences, email mapping, chair markers, science tags, and tag priority order.

    Appending `*` to any part of a member's name marks them as a chair who should receive leftover assignments.
    Tags are listed left-to-right in priority order: the first tag listed has the highest priority (index 0).
    """
    try:
        content = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise FileNotFoundError(f"Unable to read PC members file: {path}") from exc

    members: List[str] = []
    fixed: Dict[str, Dict[str, List[str]]] = {}
    emails: Dict[str, str] = {}
    chairs: Set[str] = set()
    member_tags: Dict[str, Set[str]] = {}
    member_tag_priority: Dict[str, Dict[str, int]] = {}  # member → {tag: priority_index}
    for raw_line in content.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        tokens = line.split()
        name_tokens: List[str] = []
        first_fixed: List[str] = []
        second_fixed: List[str] = []
        email: Optional[str] = None
        is_chair = False
        tags: Set[str] = set()
        tag_order: Dict[str, int] = {}
        for token in tokens:
            if "#" in token:
                try:
                    proposal_code, slot = token.split("#", 1)
                except ValueError:
                    continue
                proposal_code = proposal_code.strip()
                if not proposal_code:
                    continue
                if slot == "1":
                    first_fixed.append(proposal_code)
                elif slot == "2":
                    second_fixed.append(proposal_code)
            elif email is None and "@" in token:
                cleaned_email = token.strip("<>[](){};,")
                if cleaned_email:
                    email = cleaned_email
            else:
                chair_token = "*" in token
                cleaned_token = token.replace("*", "")
                normalized_for_alias = re.sub(r"[^\w\s-]", "", cleaned_token)
                lower_cleaned = normalized_for_alias.lower()
                if lower_cleaned in SCIENCE_CATEGORY_ALIAS_TERMS:
                    tag = SCIENCE_CATEGORY_ALIAS_TERMS[lower_cleaned]
                    if tag not in tags:
                        tag_order[tag] = len(tag_order)
                        tags.add(tag)
                    continue
                if cleaned_token:
                    name_tokens.append(cleaned_token.strip(",;"))
                if chair_token:
                    is_chair = True
        name = " ".join(name_tokens).strip()
        if not name:
            continue
        members.append(name)
        if is_chair:
            chairs.add(name)
        if tags:
            member_tags[name] = tags
            member_tag_priority[name] = tag_order
        if first_fixed or second_fixed:
            fixed[name] = {
                "first": first_fixed,
                "second": second_fixed,
            }
        if email:
            emails[name] = email

    if not members:
        raise ValueError(f"No PC members found in {path}")
    return members, fixed, emails, chairs, member_tags, member_tag_priority


def load_conflicts_file(path: Path) -> Dict[str, Set[str]]:
    """Parse a conflicts file formatted like the reviewer_assignments appendix."""
    try:
        content = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise FileNotFoundError(f"Unable to read conflicts file: {path}") from exc

    conflicts: Dict[str, Set[str]] = defaultdict(set)
    for raw_line in content.splitlines():
        line = raw_line.strip()
        if not line or line.lower() == "conflicts:":
            continue
        if ":" not in line:
            continue
        proposal_code, names = line.split(":", 1)
        proposal_code = proposal_code.strip()
        if not proposal_code:
            continue
        entries = [entry.strip() for entry in names.split(",") if entry.strip()]
        if not entries or (len(entries) == 1 and entries[0].lower() == "none"):
            continue
        for entry in entries:
            conflicts[proposal_code].add(entry)
    return conflicts


def assign_reviewers(
    proposals: List[Dict[str, Any]],
    members: Sequence[str],
    reviewers_per_proposal: int,
    max_per_member: Optional[int] = None,
    fixed_preferences: Optional[Dict[str, Dict[str, List[str]]]] = None,
    max_first_per_member: Optional[int] = None,
    max_second_per_member: Optional[int] = None,
    chair_members: Optional[Set[str]] = None,
    manual_conflicts: Optional[Dict[str, Set[str]]] = None,
    prefer_science_tags: bool = False,
    member_science_tags: Optional[Dict[str, Set[str]]] = None,
    member_tag_priority: Optional[Dict[str, Dict[str, int]]] = None,
) -> Dict[str, List[Tuple[str, str]]]:
    """Assign reviewers while balancing load, respecting per-role limits, and prioritising chairs/preferred expertise."""
    if not members:
        raise ValueError("Cannot assign reviewers without PC members.")
    if reviewers_per_proposal < 2:
        raise ValueError("Each proposal must have at least two reviewers.")
    if max_per_member is not None and max_per_member <= 0:
        raise ValueError("Maximum proposals per member must be positive.")
    if max_first_per_member is not None and max_first_per_member <= 0:
        raise ValueError("Maximum primary reviewer assignments per member must be positive.")
    if max_second_per_member is not None and max_second_per_member <= 0:
        raise ValueError("Maximum secondary reviewer assignments per member must be positive.")
    if reviewers_per_proposal > len(members):
        raise ValueError("Not enough PC members to satisfy reviewers-per-proposal.")
    chair_members = chair_members or set()
    member_science_tags = member_science_tags or {}
    member_tag_priority = member_tag_priority or {}
    member_infos = [
        {
            "name": name,
            "normalised": normalise_name(name),
            "count": 0,
            "first_count": 0,
            "second_count": 0,
            "is_chair": name in chair_members,
            "science_tags": set(member_science_tags.get(name, set())),
            "tag_priority": dict(member_tag_priority.get(name, {})),
            "order": idx,
        }
        for idx, name in enumerate(members)
        if normalise_name(name)
    ]
    if not member_infos:
        raise ValueError("PC member list does not contain valid names.")
    per_member: Dict[str, List[Tuple[str, str]]] = {info["name"]: [] for info in member_infos}
    members_by_name: Dict[str, dict] = {info["name"]: info for info in member_infos}
    members_by_normalised: Dict[str, str] = {
        info["normalised"]: info["name"] for info in member_infos if info["normalised"]
    }
    role_labels = generate_role_labels(reviewers_per_proposal)

    fixed_first_map: Dict[str, str] = {}
    fixed_second_map: Dict[str, str] = {}
    fixed_assignments: Set[Tuple[str, str]] = set()
    if fixed_preferences:
        for member_name, slots in fixed_preferences.items():
            if member_name not in members_by_name:
                raise ValueError(f"PC member '{member_name}' with fixed assignments is not in the member list.")
            for code in slots.get("first", []):
                if not code:
                    continue
                existing = fixed_first_map.get(code)
                if existing and existing != member_name:
                    raise ValueError(f"Conflicting primary reviewer assignment for proposal {code}.")
                fixed_first_map[code] = member_name
            for code in slots.get("second", []):
                if not code:
                    continue
                existing = fixed_second_map.get(code)
                if existing and existing != member_name:
                    raise ValueError(f"Conflicting secondary reviewer assignment for proposal {code}.")
                fixed_second_map[code] = member_name

    PRIMARY_ROLE = "Primary Reviewer"
    SECONDARY_ROLE = "Secondary Reviewer"

    # Pre-assign primary reviewer for proposals where a member is the best expert based on
    # unique tag ownership.  "Best" = most unique-tag matches for the proposal; ties broken by
    # the expert's tag priority (first-listed tag wins).  This ensures that e.g. a member listed
    # as "Maser SpectralLine" gets Maser proposals assigned before generic SpectralLine ones.
    # reserved_primaries tracks how many primary slots to protect from load-balancing.
    reserved_primaries: Dict[str, int] = {}
    if prefer_science_tags and max_first_per_member is not None:
        # Map each tag to the members who hold it.
        _tag_holders: Dict[str, List[dict]] = {}
        for _m in member_infos:
            for _t in _m["science_tags"]:
                _tag_holders.setdefault(_t, []).append(_m)

        # For each proposal, determine the single best expert via unique-tag scoring.
        # score = (neg_unique_count, best_tag_priority) — minimise to find winner.
        expert_proposals: Dict[str, List[Tuple[int, int, str]]] = {}
        for _p_idx, _p in enumerate(proposals):
            _pcode = _p["exp"]
            if _pcode in fixed_first_map:
                continue
            _ptags = set(_p.get("science_tags") or [])
            _scores: Dict[str, Tuple[int, int]] = {}  # expert → (neg_count, best_prio)
            for _tag in _ptags:
                _holders = _tag_holders.get(_tag, [])
                if len(_holders) == 1:
                    _expert = _holders[0]
                    _ename = _expert["name"]
                    _prio = _expert.get("tag_priority", {}).get(_tag, 999)
                    _neg, _best = _scores.get(_ename, (0, 999))
                    _scores[_ename] = (_neg - 1, min(_best, _prio))
            if not _scores:
                continue
            # Pick the expert with the best score (most unique matches, then lowest priority).
            _winner = min(_scores, key=lambda e: _scores[e])
            _best_prio = _scores[_winner][1]
            expert_proposals.setdefault(_winner, []).append((_best_prio, _p_idx, _pcode))

        # For each expert, sort slots by (tag_priority, proposal_index) and pre-assign up to limit.
        for _ename, _slots in expert_proposals.items():
            _seen: Dict[str, Tuple[int, int, str]] = {}
            for _entry in _slots:
                _pcode = _entry[2]
                if _pcode not in _seen or _entry < _seen[_pcode]:
                    _seen[_pcode] = _entry
            _sorted_slots = sorted(_seen.values())

            reserved_primaries[_ename] = min(len(_sorted_slots), max_first_per_member)

            _pre_count = 0
            for _, _, _pcode in _sorted_slots:
                if _pre_count >= max_first_per_member:
                    break
                if _pcode not in fixed_first_map:
                    fixed_first_map[_pcode] = _ename
                    _pre_count += 1

    def has_capacity(member: dict, role: str) -> bool:
        if role == PRIMARY_ROLE and max_first_per_member is not None:
            if member["first_count"] >= max_first_per_member:
                return False
        if role == SECONDARY_ROLE and max_second_per_member is not None:
            if member["second_count"] >= max_second_per_member:
                return False
        return True

    def record_assignment(member: dict, role: str) -> None:
        member["count"] += 1
        if role == PRIMARY_ROLE:
            member["first_count"] += 1
        elif role == SECONDARY_ROLE:
            member["second_count"] += 1

    def remove_assignment(member: dict, role: str) -> None:
        member["count"] = max(0, member["count"] - 1)
        if role == PRIMARY_ROLE:
            member["first_count"] = max(0, member["first_count"] - 1)
        elif role == SECONDARY_ROLE:
            member["second_count"] = max(0, member["second_count"] - 1)

    def priority_key(member: dict, role: str) -> Tuple[int, int, int, int]:
        """Return a tuple used to balance role-specific assignments (chairs soak up leftovers)."""
        chair_bias = 0 if member.get("is_chair") else 1
        if role == PRIMARY_ROLE:
            return (member["first_count"], chair_bias, member["count"], member["order"])
        if role == SECONDARY_ROLE:
            return (member["second_count"], chair_bias, member["count"], member["order"])
        return (member["count"], chair_bias, member["order"], 0)

    def select_member(excluded: Set[str], already_chosen: Set[str], role: str, proposal_tags: Set[str]) -> dict:
        eligible = [
            member
            for member in member_infos
            if member["normalised"] not in excluded
            and member["normalised"] not in already_chosen
            and (max_per_member is None or member["count"] < max_per_member)
            and has_capacity(member, role)
        ]
        if not eligible:
            eligible = [
                member
                for member in member_infos
                if member["normalised"] not in excluded
                and (max_per_member is None or member["count"] < max_per_member)
                and has_capacity(member, role)
            ]
        if not eligible:
            raise ValueError("No available reviewers remaining for assignment within limits.")
        if prefer_science_tags and role == PRIMARY_ROLE:
            if proposal_tags:
                preferred = [
                    member for member in eligible if member["science_tags"] and member["science_tags"].intersection(proposal_tags)
                ]
                if preferred:
                    eligible = preferred
            if not (proposal_tags and any(m["science_tags"] & proposal_tags for m in eligible)):
                # No tag match: protect reserved primary slots so sole-experts keep
                # capacity for their pre-assigned proposals.
                if reserved_primaries and max_first_per_member is not None:
                    non_reserved = [
                        m for m in eligible
                        if (max_first_per_member - m["first_count"]) > reserved_primaries.get(m["name"], 0)
                    ]
                    if non_reserved:
                        eligible = non_reserved
        chosen = min(eligible, key=lambda m: priority_key(m, role))
        return chosen

    def apply_fixed_member(
        member_name: str,
        role_idx: int,
        proposal_code: str,
        excluded: Set[str],
        already_chosen: Set[str],
        reviewers: List[Tuple[str, str]],
    ) -> None:
        member = members_by_name.get(member_name)
        if member is None:
            raise ValueError(f"Fixed reviewer '{member_name}' is not a recognised PC member.")
        if member["normalised"] in excluded:
            raise ValueError(
                f"Fixed reviewer '{member_name}' is listed on proposal {proposal_code}, cannot assign."
            )
        if member["normalised"] in already_chosen:
            raise ValueError(
                f"Fixed reviewer '{member_name}' requested multiple slots for proposal {proposal_code}."
            )
        if max_per_member is not None and member["count"] >= max_per_member:
            raise ValueError(
                f"Fixed reviewer '{member_name}' exceeds the maximum assignments ({max_per_member})."
            )
        role = role_labels[role_idx]
        if not has_capacity(member, role):
            limit_label = "primary reviewer" if role == PRIMARY_ROLE else "secondary reviewer"
            raise ValueError(
                f"Fixed reviewer '{member_name}' exceeds the maximum {limit_label} assignments."
            )
        record_assignment(member, role)
        already_chosen.add(member["normalised"])
        reviewers.append((role, member_name))
        per_member[member_name].append((proposal_code, role))
        fixed_assignments.add((proposal_code, role))

    for proposal in proposals:
        proposal_code = proposal["exp"]
        proposal_tags: Set[str] = set(proposal.get("science_tags") or [])
        participants: Set[str] = set(proposal.get("participants", set()))
        excluded = set(participants)
        # Capture everyone we exclude so the CSV can describe conflicts explicitly.
        conflicts: Set[str] = set()
        if participants:
            for member in member_infos:
                norm = member["normalised"]
                if norm and norm in participants:
                    conflicts.add(member["name"])
        text_blob = proposal.get("normalised_text", "")
        if text_blob:
            for member in member_infos:
                norm = member["normalised"]
                if norm and f" {norm} " in text_blob:
                    excluded.add(norm)
                    conflicts.add(member["name"])
        if manual_conflicts:
            extra_conflicts = manual_conflicts.get(proposal_code, set())
            for entry in extra_conflicts:
                resolved_name = entry.strip()
                norm = normalise_name(resolved_name)
                if norm:
                    excluded.add(norm)
                    resolved_name = members_by_normalised.get(norm, resolved_name)
                if resolved_name:
                    conflicts.add(resolved_name)
        chosen: Set[str] = set()
        reviewers: List[Tuple[str, str]] = []
        fixed_slots: List[Optional[str]] = [None] * reviewers_per_proposal
        fixed_first = fixed_first_map.get(proposal_code)
        fixed_second = fixed_second_map.get(proposal_code)
        if fixed_first:
            fixed_slots[0] = fixed_first
        if fixed_second:
            if reviewers_per_proposal < 2:
                raise ValueError(
                    f"Cannot assign a secondary reviewer for {proposal_code} when reviewers-per-proposal < 2."
                )
            if fixed_first and fixed_second == fixed_first:
                raise ValueError(
                    f"Member '{fixed_first}' cannot be both primary and secondary reviewer for {proposal_code}."
                )
            fixed_slots[1] = fixed_second

        for idx in range(reviewers_per_proposal):
            fixed_member = fixed_slots[idx]
            role = role_labels[idx]
            if fixed_member:
                apply_fixed_member(fixed_member, idx, proposal_code, excluded, chosen, reviewers)
                continue
            member = select_member(excluded, chosen, role, proposal_tags)
            chosen.add(member["normalised"])
            record_assignment(member, role)
            reviewers.append((role, member["name"]))
            per_member[member["name"]].append((proposal_code, role))

        proposal["reviewers"] = reviewers
        proposal["conflicts"] = sorted(conflicts)

    proposal_lookup: Dict[str, Dict[str, Any]] = {proposal["exp"]: proposal for proposal in proposals}

    def rebalance_role_assignments(role_label: str) -> None:
        """Reassign slots so every member stays within one assignment of each other for the role."""
        if role_label == PRIMARY_ROLE:
            count_key = "first_count"
        elif role_label == SECONDARY_ROLE:
            count_key = "second_count"
        else:
            return

        blocked_donors: Set[str] = set()

        def can_receive(member: dict) -> bool:
            if max_per_member is not None and member["count"] >= max_per_member:
                return False
            return has_capacity(member, role_label)

        def find_transfer_slot(donor_info: dict, recipient_info: dict) -> Optional[Tuple[Dict[str, Any], int]]:
            if not can_receive(recipient_info):
                return None
            donor_name = donor_info["name"]
            recipient_name = recipient_info["name"]
            recipient_norm = recipient_info.get("normalised")
            assignments = per_member.get(donor_name, [])
            for idx, (proposal_code, role) in enumerate(assignments):
                if role != role_label:
                    continue
                if (proposal_code, role_label) in fixed_assignments:
                    continue
                proposal = proposal_lookup.get(proposal_code)
                if not proposal:
                    continue
                if any(name == recipient_name for _role, name in proposal.get("reviewers", [])):
                    continue
                excluded_norms: Set[str] = set(proposal.get("participants", set()))
                for entry in proposal.get("conflicts") or []:
                    norm = normalise_name(entry)
                    if norm:
                        excluded_norms.add(norm)
                if recipient_norm and recipient_norm in excluded_norms:
                    continue
                # Don't rebalance a tag-matched assignment to a member with no matching expertise.
                if prefer_science_tags:
                    ptags = set(proposal.get("science_tags") or [])
                    if ptags:
                        donor_tags = donor_info.get("science_tags", set())
                        recipient_tags = recipient_info.get("science_tags", set())
                        if donor_tags & ptags and not (recipient_tags & ptags):
                            continue
                return proposal, idx
            return None

        while True:
            if not member_infos:
                break
            counts = [member[count_key] for member in member_infos]
            max_count = max(counts)
            min_count = min(counts)
            if max_count - min_count <= 1:
                break

            donors = [
                info
                for info in member_infos
                if info[count_key] == max_count and info["name"] not in blocked_donors
            ]
            recipients = [
                info
                for info in member_infos
                if info[count_key] == min_count and can_receive(info)
            ]
            if not donors or not recipients:
                break

            donors.sort(key=lambda info: (-info[count_key], info["order"]))
            recipients.sort(key=lambda info: (info[count_key], info["order"]))
            transfer_made = False

            for donor in donors:
                for recipient in recipients:
                    if donor["name"] == recipient["name"]:
                        continue
                    slot = find_transfer_slot(donor, recipient)
                    if not slot:
                        continue
                    proposal, donor_idx = slot
                    proposal_code = proposal["exp"]

                    for idx, (role, name) in enumerate(proposal["reviewers"]):
                        if role == role_label and name == donor["name"]:
                            proposal["reviewers"][idx] = (role_label, recipient["name"])
                            break

                    donor_assignments = per_member.get(donor["name"], [])
                    if donor_idx < len(donor_assignments):
                        donor_assignments.pop(donor_idx)
                    per_member[recipient["name"]].append((proposal_code, role_label))

                    remove_assignment(donor, role_label)
                    record_assignment(recipient, role_label)
                    transfer_made = True
                    blocked_donors.clear()
                    break
                if transfer_made:
                    break
                blocked_donors.add(donor["name"])

            if not transfer_made:
                break

    rebalance_role_assignments(PRIMARY_ROLE)
    rebalance_role_assignments(SECONDARY_ROLE)

    for proposal in proposals:
        reviewers = proposal.get("reviewers", [])
        proposal["first_reviewer"] = next(
            (name for role, name in reviewers if role == PRIMARY_ROLE),
            None,
        )
        proposal["second_reviewer"] = next(
            (name for role, name in reviewers if role == SECONDARY_ROLE),
            None,
        )
        additional = [name for role, name in reviewers if role not in {PRIMARY_ROLE, SECONDARY_ROLE}]
        if additional:
            proposal["additional_reviewers"] = additional
        else:
            proposal.pop("additional_reviewers", None)

    return per_member


def write_assignments(proposals: Sequence[Dict[str, Any]], destination: Path, roles: Sequence[str]) -> None:
    """Persist reviewer assignments to a CSV file and append a conflicts appendix."""
    roles = list(roles)
    max_count = max((len(proposal.get("reviewers", [])) for proposal in proposals), default=0)
    if max_count > len(roles):
        extra_roles = generate_role_labels(max_count)[len(roles):]
        roles.extend(extra_roles)

    header = ["Proposal", *roles]
    csv_lines = [",".join(header)]

    for proposal in proposals:
        row = [proposal["exp"]]
        reviewers = proposal.get("reviewers", [])
        role_map = {role: name for role, name in reviewers}
        for role in roles:
            row.append(role_map.get(role, ""))
        csv_lines.append(",".join(row))

    conflict_lines: List[str] = []
    for proposal in proposals:
        conflicts = proposal.get("conflicts") or []
        if conflicts:
            conflict_lines.append(f"{proposal['exp']}: {', '.join(conflicts)}")
        else:
            conflict_lines.append(f"{proposal['exp']}: None")

    destination.parent.mkdir(parents=True, exist_ok=True)
    sections = ["\n".join(csv_lines)]
    ascii_table = build_role_ascii_table(proposals, roles)
    if ascii_table:
        sections.append("Reviewer Summary:")
        sections.append(ascii_table)
    if conflict_lines:
        sections.append("Conflicts:")
        sections.append("\n".join(conflict_lines))
    content = "\n\n".join(sections)
    if not content.endswith("\n"):
        content += "\n"
    destination.write_text(content, encoding="utf-8")


def write_science_tags(proposals: Sequence[Dict[str, Any]], destination: Path) -> None:
    """Write inferred science categories per proposal."""
    lines: List[str] = []
    for proposal in proposals:
        tags = proposal.get("science_tags") or [SCIENCE_DEFAULT_CATEGORY]
        lines.append(f"{proposal['exp']}: {', '.join(tags)}")
    content = "\n".join(lines)
    if not content.endswith("\n"):
        content += "\n"
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(content, encoding="utf-8")


def read_science_tags_file(source: Path) -> Dict[str, List[str]]:
    """Read a science-tags file and return a mapping of proposal code → tag list."""
    overrides: Dict[str, List[str]] = {}
    for raw_line in source.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if ":" not in line:
            continue
        code, _, tag_str = line.partition(":")
        tags = [t.strip() for t in tag_str.split(",") if t.strip()]
        if code.strip() and tags:
            overrides[code.strip()] = tags
    return overrides


def write_pi_emails(proposals: Sequence[Dict[str, Any]], destination: Path) -> None:
    """Write PI email addresses per proposal for post-assessment feedback."""
    lines: List[str] = []
    for proposal in proposals:
        email = proposal.get("pi_email") or ""
        lines.append(f"{proposal['exp']}: {email}")
    content = "\n".join(lines)
    if not content.endswith("\n"):
        content += "\n"
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(content, encoding="utf-8")


DOCX_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
WORDPROCESSING_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
DOCX_COMMENT_AUTHOR = "EVN Programme Committee"
DOCX_SEPARATOR = "=" * 69

ROOT_RELATIONSHIPS_XML = (
    f"{DOCX_XML_DECL}\n"
    f'<Relationships xmlns="{PACKAGE_REL_NS}">\n'
    f'  <Relationship Id="rId1" Type="{DOCUMENT_REL_NS}/officeDocument" Target="word/document.xml"/>\n'
    "</Relationships>"
)

DEFAULT_STYLES_XML = (
    f"{DOCX_XML_DECL}\n"
    f'<w:styles xmlns:w="{WORDPROCESSING_NS}">\n'
    '  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">\n'
    '    <w:name w:val="Normal"/>\n'
    "    <w:qFormat/>\n"
    "  </w:style>\n"
    '  <w:style w:type="paragraph" w:styleId="Title">\n'
    '    <w:name w:val="Title"/>\n'
    '    <w:basedOn w:val="Normal"/>\n'
    "    <w:qFormat/>\n"
    "    <w:rPr>\n"
    "      <w:b/>\n"
    '      <w:sz w:val="56"/>\n'
    "    </w:rPr>\n"
    "  </w:style>\n"
    '  <w:style w:type="paragraph" w:styleId="Heading1">\n'
    '    <w:name w:val="Heading 1"/>\n'
    '    <w:basedOn w:val="Normal"/>\n'
    "    <w:qFormat/>\n"
    "    <w:rPr>\n"
    "      <w:b/>\n"
    '      <w:sz w:val="32"/>\n'
    "    </w:rPr>\n"
    "  </w:style>\n"
    '  <w:style w:type="paragraph" w:styleId="Heading2">\n'
    '    <w:name w:val="Heading 2"/>\n'
    '    <w:basedOn w:val="Normal"/>\n'
    "    <w:qFormat/>\n"
    "    <w:rPr>\n"
    "      <w:b/>\n"
    '      <w:sz w:val="28"/>\n'
    "    </w:rPr>\n"
    "  </w:style>\n"
    "</w:styles>"
)


def needs_space_preservation(text: str) -> bool:
    return bool(text) and (text.startswith(" ") or text.endswith(" ") or "  " in text)


def paragraph_xml(text: str, comment_id: Optional[int], style: Optional[str] = None) -> str:
    indent = "    "
    if not text:
        return f"{indent}<w:p/>"
    attr = ' xml:space="preserve"' if needs_space_preservation(text) else ""
    ppr = f"<w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr>" if style else ""
    text_xml = f"<w:r><w:t{attr}>{escape(text)}</w:t></w:r>"
    if comment_id is None:
        return f"{indent}<w:p>{ppr}{text_xml}</w:p>"
    return (
        f'{indent}<w:p>{ppr}<w:commentRangeStart w:id="{comment_id}"/>{text_xml}'
        f'<w:commentRangeEnd w:id="{comment_id}"/><w:r><w:commentReference w:id="{comment_id}"/></w:r></w:p>'
    )


def comment_paragraph_xml(text: str) -> str:
    if not text:
        return "    <w:p/>"
    attr = ' xml:space="preserve"' if needs_space_preservation(text) else ""
    return f'    <w:p><w:r><w:t{attr}>{escape(text)}</w:t></w:r></w:p>'


def build_comments_xml(entries: Sequence[Tuple[int, str]]) -> str:
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    lines = [
        DOCX_XML_DECL,
        f'<w:comments xmlns:w="{WORDPROCESSING_NS}">',
    ]
    for comment_id, text in entries:
        lines.append(
            f'  <w:comment w:id="{comment_id}" w:author="{DOCX_COMMENT_AUTHOR}" w:date="{timestamp}">'
        )
        comment_lines = text.splitlines() or [""]
        for chunk in comment_lines:
            lines.append(comment_paragraph_xml(chunk))
        lines.append("  </w:comment>")
    lines.append("</w:comments>")
    return "\n".join(lines)


DocxLine = Union[str, Tuple[str, Optional[str]], Dict[str, Any]]


def extract_docx_line(entry: DocxLine) -> Tuple[str, Optional[str]]:
    if isinstance(entry, dict):
        text = str(entry.get("text", ""))
        style = entry.get("style")
        return text, style
    if isinstance(entry, tuple):
        if not entry:
            return "", None
        text = str(entry[0])
        style = entry[1] if len(entry) > 1 else None
        return text, style
    return str(entry), None


def build_document_xml(records: Sequence[DocxRecord]) -> Tuple[str, Optional[str]]:
    lines = [
        DOCX_XML_DECL,
        f'<w:document xmlns:w="{WORDPROCESSING_NS}">',
        "  <w:body>",
    ]
    comment_entries: List[Tuple[int, str]] = []
    next_comment_id = 0

    for record_idx, record in enumerate(records):
        comment_id: Optional[int] = None
        comment_text = record.get("comment")
        if comment_text:
            comment_id = next_comment_id
            comment_entries.append((comment_id, comment_text))
            next_comment_id += 1
        record_lines: Sequence[DocxLine] = record.get("lines") or []
        for line_idx, entry in enumerate(record_lines):
            text, style = extract_docx_line(entry)
            target_comment_id = comment_id if line_idx == 0 else None
            lines.append(paragraph_xml(text, target_comment_id, style))
        if record_idx < len(records) - 1:
            lines.append('    <w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    lines.append("    <w:sectPr>")
    lines.append('      <w:pgSz w:w="12240" w:h="15840"/>')
    lines.append(
        '      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="708" w:footer="708" w:gutter="0"/>'
    )
    lines.append("    </w:sectPr>")
    lines.append("  </w:body>")
    lines.append("</w:document>")

    document_xml = "\n".join(lines)
    comments_xml = build_comments_xml(comment_entries) if comment_entries else None
    return document_xml, comments_xml


def build_document_relationships_xml(include_comments: bool) -> str:
    lines = [
        DOCX_XML_DECL,
        f'<Relationships xmlns="{PACKAGE_REL_NS}">',
        f'  <Relationship Id="rId1" Type="{DOCUMENT_REL_NS}/styles" Target="styles.xml"/>',
    ]
    if include_comments:
        lines.append(
            f'  <Relationship Id="rId2" Type="{DOCUMENT_REL_NS}/comments" Target="comments.xml"/>'
        )
    lines.append("</Relationships>")
    return "\n".join(lines)


def build_content_types_xml(include_comments: bool) -> str:
    lines = [
        DOCX_XML_DECL,
        f'<Types xmlns="{CONTENT_TYPES_NS}">',
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '  <Default Extension="xml" ContentType="application/xml"/>',
        '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
        '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    ]
    if include_comments:
        lines.append(
            '  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
        )
    lines.append("</Types>")
    return "\n".join(lines)


def build_comment_text(
    first_reviewer: Optional[str],
    second_reviewer: Optional[str],
) -> Optional[str]:
    entries: List[str] = []
    if first_reviewer:
        entries.append(f"Primary reviewer: {first_reviewer}")
    if second_reviewer:
        entries.append(f"Secondary reviewer: {second_reviewer}")
    if not entries:
        return None
    return "\n".join(entries)


def write_docx_records(records: Sequence[DocxRecord], destination: Path) -> None:
    if not records:
        log_verbose("No DOCX records to write; skipping file creation.")
        return
    log_verbose(f"Preparing DOCX document with {len(records)} section(s).")
    document_xml, comments_xml = build_document_xml(records)
    include_comments = comments_xml is not None
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destination, "w") as archive:
        archive.writestr("[Content_Types].xml", build_content_types_xml(include_comments))
        archive.writestr("_rels/.rels", ROOT_RELATIONSHIPS_XML)
        archive.writestr("word/_rels/document.xml.rels", build_document_relationships_xml(include_comments))
        archive.writestr("word/styles.xml", DEFAULT_STYLES_XML)
        archive.writestr("word/document.xml", document_xml)
        if include_comments and comments_xml:
            archive.writestr("word/comments.xml", comments_xml)


def build_agenda_lines() -> List[DocxLine]:
    return [
        {"text": "EVN PC meeting __________________ (____________________)", "style": "Title"},
        {"text": "Agenda v1.0", "style": "Heading2"},
        "",
        "Meeting dates (edit as needed): ________________________________",
        "Meeting location (city, country): ______________________________",
        "",
        "Join Zoom Meeting: ________________________________",
        "",
        "Start of meeting: ________________________________",
        "",
        {"text": "1. Welcome, approval of the agenda – JR – 5'", "style": "Heading1"},
        "",
        {"text": "2. Minutes of the last meeting, action items – 10'", "style": "Heading1"},
        "",
        {"text": "3. Updates – brief, roundtable – 20'", "style": "Heading1"},
        {"text": "3.1 Station updates", "style": "Heading2"},
        {"text": "3.2 Correlator updates – RMC, AL", "style": "Heading2"},
        {"text": "3.3 Scheduler's report, EVN resources", "style": "Heading2"},
        {"text": "3.4 Review of PC membership", "style": "Heading2"},
        {"text": "3.5 EVN PC issues/decisions since the last meeting", "style": "Heading2"},
        "(see list at the end of the Agenda)",
        "",
        {"text": "4. Proposals of the 2024C call", "style": "Heading1"},
        "",
        {"text": "5. Scheduler's summary – 10'", "style": "Heading1"},
        "Review of decisions and impact on sessions.",
        "Future prospects.",
        "",
        {"text": "6. AOB", "style": "Heading1"},
        "",
        "End at ________________________________",
        "",
        {"text": "EVN PC Chair decisions since the last meeting", "style": "Heading1"},
        "______________________________________________",
        "",
        {"text": "Notes:", "style": "Heading2"},
        "______________________________________________",
    ]


def build_role_ascii_table(proposals: Sequence[Dict[str, Any]], roles: Sequence[str]) -> str:
    """Return an ASCII table listing each proposal and reviewer per role."""
    roles = list(roles)
    if not proposals or not roles:
        return ""

    headers = ["Proposal", *roles]
    rows: List[List[str]] = []
    for proposal in proposals:
        role_map = {role: reviewer for role, reviewer in proposal.get("reviewers", [])}
        row = [proposal["exp"], *[role_map.get(role, "") for role in roles]]
        rows.append(row)

    col_widths = [len(header) for header in headers]
    for row in rows:
        for idx, cell in enumerate(row):
            col_widths[idx] = max(col_widths[idx], len(cell))

    def fmt_row(values: List[str]) -> str:
        parts = [
            f" {value.ljust(col_widths[idx])} "
            for idx, value in enumerate(values)
        ]
        return "|" + "|".join(parts) + "|"

    border = "+" + "+".join("-" * (width + 2) for width in col_widths) + "+"
    lines = [border, fmt_row(headers), border]
    for row in rows:
        lines.append(fmt_row(row))
    lines.append(border)
    return "\n".join(lines)


def build_reviewer_email_table(
    assignments: Dict[str, List[Tuple[str, str]]],
    roles: Sequence[str],
    member_tags: Optional[Dict[str, Set[str]]] = None,
) -> str:
    """Return an HTML table summarizing per-reviewer assignments by role."""
    if not assignments:
        return ""

    base_role_display = {
        "Primary Reviewer": "Primary",
        "Secondary Reviewer": "Secondary",
    }
    base_roles: List[Tuple[str, str]] = []
    extra_roles: List[str] = []
    for role in roles:
        lower = role.lower()
        if lower.startswith("additional review"):
            extra_roles.append(role)
        elif role in base_role_display:
            base_roles.append((role, base_role_display[role]))

    headers = ["Reviewer", *[display for _, display in base_roles]]
    if extra_roles:
        headers.append("Additional")

    lines = [
        '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;">',
        "  <thead>",
        "    <tr>"
        + "".join(f"<th>{html.escape(header)}</th>" for header in headers)
        + "</tr>",
        "  </thead>",
        "  <tbody>",
    ]

    for reviewer in sorted(assignments.keys(), key=str.lower):
        row_cells: List[str] = [html.escape(reviewer)]
        slot_map: Dict[str, List[str]] = defaultdict(list)
        for proposal_code, role in assignments[reviewer]:
            slot_map[role].append(proposal_code)
        for role, _display in base_roles:
            entries = slot_map.get(role)
            row_cells.append(format_assignment_entries(entries))
        if extra_roles:
            combined: List[str] = []
            for role in extra_roles:
                combined.extend(slot_map.get(role, []))
            row_cells.append(format_assignment_entries(combined))
        lines.append("    <tr>" + "".join(f"<td>{cell}</td>" for cell in row_cells) + "</tr>")

    lines.append("  </tbody>")
    lines.append("</table>")

    return "\n".join(lines)


def format_assignment_entries(entries: Optional[Sequence[str]]) -> str:
    """Format assignment codes for a table cell."""
    if not entries:
        return ""
    # Sort proposal codes alphabetically within each role column for readability.
    sorted_entries = sorted(entries, key=str.lower)
    return ", ".join(html.escape(entry) for entry in sorted_entries)


def write_member_summary(
    assignments: Dict[str, List[Tuple[str, str]]],
    destination: Path,
    roles: Sequence[str],
    member_tags: Optional[Dict[str, Set[str]]] = None,
) -> None:
    """Persist per-member assignments as an HTML table for easy emailing."""
    table_html = build_reviewer_email_table(assignments, roles, member_tags)
    if not table_html:
        table_html = "<p>No reviewer assignments available.</p>"

    destination.parent.mkdir(parents=True, exist_ok=True)
    html_content = "\n".join(
        [
            "<!-- Reviewer assignments for Outlook. Paste the table below into the email body. -->",
            table_html,
        ]
    )
    if not html_content.endswith("\n"):
        html_content += "\n"
    destination.write_text(html_content, encoding="utf-8")


def parse_proposal(path: Path) -> Dict[str, Any]:
    """Derive structured proposal metadata from a PDF."""
    lines = extract_pdf_lines(path)
    exp, title, pi_hint = find_experiment_and_title(lines, path.stem)
    normalised_text = f" {normalise_name(' '.join(lines))} "

    applicants = parse_applicants(lines)
    pi_row = next(
        (row for row in applicants if any(
            v.strip().lower() in {"pi", "p.i."}
            for v in row.values()
        )),
        None,
    )
    if pi_row is None and applicants:
        pi_row = applicants[0]
    pi = pi_row.get("name") if pi_row else None
    pi_email_raw = pi_row.get("email", "").strip() if pi_row else ""
    # The Contact Author section has labelled fields and is more reliably parsed
    # than the fixed-width Applicants table (where column overlap can corrupt the
    # email field).  Prefer it; fall back to the applicant-table value.
    pi_email: Optional[str] = parse_contact_email(lines) or _best_email(pi_email_raw)
    if not pi:
        pi = pi_hint or parse_contact_name(lines) or "Unknown"

    participants: Set[str] = set()
    for row in applicants:
        name = row.get("name", "")
        normalised = normalise_name(name)
        if normalised:
            participants.add(normalised)

    contact_name = parse_contact_name(lines)
    if contact_name:
        normalised_contact = normalise_name(contact_name)
        if normalised_contact:
            participants.add(normalised_contact)
    normalised_pi = normalise_name(pi)
    if normalised_pi:
        participants.add(normalised_pi)

    networks, wavebands = parse_summary(lines)
    nets = ", ".join(networks)
    lambdas = ", ".join(wavebands)
    science_tags = infer_science_categories(title, lines)

    return {
        "exp": exp,
        "pi": pi,
        "pi_email": pi_email,
        "title": title,
        "nets": nets,
        "lambda": lambdas,
        "participants": participants,
        "normalised_text": normalised_text,
        "science_tags": science_tags,
    }


def iter_pdf_paths(supplied: Sequence[Path], pdf_dir: Optional[Path]) -> Iterable[Path]:
    """Yield the PDFs to process, combining explicit files with directory listings."""
    if pdf_dir is not None and not pdf_dir.is_dir():
        raise FileNotFoundError(f"PDF directory not found: {pdf_dir}")

    if supplied:
        for supplied_path in supplied:
            if supplied_path.is_file():
                yield supplied_path
                continue
            if pdf_dir is not None:
                candidate = supplied_path if supplied_path.is_absolute() else pdf_dir / supplied_path
                if candidate.is_file():
                    yield candidate
                    continue
            raise FileNotFoundError(f"PDF not found: {supplied_path}")
    else:
        if pdf_dir is None:
            raise FileNotFoundError("No PDF directory provided and no PDF files supplied.")
        for path in sorted(pdf_dir.glob("*.pdf")):
            if path.is_file():
                yield path


def main(argv: Optional[Sequence[str]] = None) -> int:
    """CLI entry point for rendering templates and assigning reviewers."""
    parser = argparse.ArgumentParser(
        description="Render EVN proposal PDFs in chair template format.",
        add_help=False,
    )
    parser.add_argument(
        "-h",
        "--help",
        action="help",
        help="Show this help message and exit.",
    )
    parser.add_argument(
        "pdfs",
        nargs="*",
        type=Path,
        help="Specific PDF files to process. Defaults to all PDFs in --pdf-dir.",
    )
    parser.add_argument(
        "-p",
        "--pdf-dir",
        type=Path,
        help="Directory containing proposal PDFs (required if no PDF files are listed).",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Write the formatted output to this file instead of stdout.",
    )
    parser.add_argument(
        "-A",
        "--agenda-docx",
        type=Path,
        help="Write the meeting agenda DOCX to this file.",
    )
    parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="Suppress progress information on stderr.",
    )
    parser.add_argument(
        "-m",
        "--pc-members",
        type=Path,
        help="File containing EVN PC members (`Name Email` per line, optional CODE#slot) to auto-assign reviewers.",
    )
    parser.add_argument(
        "-a",
        "--assignments",
        type=Path,
        help="File to store reviewer assignments (defaults to reviewer_assignments.txt when --pc-members is used).",
    )
    parser.add_argument(
        "-R",
        "--reviewers-per-proposal",
        type=int,
        default=2,
        metavar="COUNT",
        help="Total reviewers to assign per proposal (minimum 2, default 2).",
    )
    parser.add_argument(
        "-M",
        "--max-per-member",
        type=int,
        metavar="MAX",
        help="Maximum number of proposals assigned to any single PC member.",
    )
    parser.add_argument(
        "-F",
        "--max-first-per-member",
        type=int,
        metavar="COUNT",
        help="Maximum number of primary reviewer assignments per PC member.",
    )
    parser.add_argument(
        "-S",
        "--max-second-per-member",
        type=int,
        metavar="COUNT",
        help="Maximum number of secondary reviewer assignments per PC member.",
    )
    parser.add_argument(
        "-H",
        "--member-summary",
        type=Path,
        help="Write a per-member HTML assignment table to this file.",
    )
    parser.add_argument(
        "-c",
        "--conflicts-file",
        type=Path,
        help="Optional file listing per-proposal conflicts (same format as the reviewer assignment appendix).",
    )
    parser.add_argument(
        "-t",
        "--science-tags-file",
        type=Path,
        help="Write inferred science categories per proposal to this file.",
    )
    parser.add_argument(
        "-T",
        "--prefer-matching-tags",
        action="store_true",
        help="Prefer matching primary reviewers whose expertise tags overlap the proposal science tags.",
    )
    parser.add_argument(
        "--pi-emails-file",
        type=Path,
        help="Write PI email addresses per proposal to this file (for post-assessment feedback).",
    )
    args = parser.parse_args(argv)
    global VERBOSE
    VERBOSE = not bool(args.quiet)

    try:
        pdf_paths = list(iter_pdf_paths(args.pdfs, args.pdf_dir))
    except FileNotFoundError as exc:
        print(exc, file=sys.stderr)
        return 1

    if not pdf_paths:
        print("No proposal PDFs found.", file=sys.stderr)
        return 1
    log_verbose(f"Discovered {len(pdf_paths)} proposal PDF(s).")

    if (
        args.assignments
        or args.max_per_member
        or args.member_summary
        or args.max_first_per_member
        or args.max_second_per_member
        or args.conflicts_file
        or args.prefer_matching_tags
    ) and not args.pc_members:
        print("Reviewer-related options require --pc-members to be specified.", file=sys.stderr)
        return 1
    if args.reviewers_per_proposal < 2:
        print("--reviewers-per-proposal must be at least 2.", file=sys.stderr)
        return 1

    proposals: List[Dict[str, Any]] = []

    for path in pdf_paths:
        try:
            proposal = parse_proposal(path)
        except PdfExtractError as exc:
            print(f"{path}: {exc}", file=sys.stderr)
            return 2
        log_verbose(f"Parsed proposal {proposal['exp']} – PI {proposal['pi']}.")
        proposal["pdf_path"] = str(path)
        proposals.append(proposal)

    # Apply manual science-tag overrides from the tags file if it already exists.
    if args.science_tags_file and args.science_tags_file.is_file():
        try:
            tag_overrides = read_science_tags_file(args.science_tags_file)
            for proposal in proposals:
                if proposal["exp"] in tag_overrides:
                    proposal["science_tags"] = tag_overrides[proposal["exp"]]
            log_verbose(f"Applied science-tag overrides from {args.science_tags_file}.")
        except OSError as exc:
            print(f"Failed to read science tags file: {exc}", file=sys.stderr)
            return 1

    member_assignments: Optional[Dict[str, List[Tuple[str, str]]]] = None
    role_labels: Optional[List[str]] = None
    manual_conflicts: Dict[str, Set[str]] = {}
    member_tags: Dict[str, Set[str]] = {}
    member_tag_priority: Dict[str, Dict[str, int]] = {}

    if args.conflicts_file:
        try:
            manual_conflicts = load_conflicts_file(args.conflicts_file)
        except (FileNotFoundError, ValueError) as exc:
            print(exc, file=sys.stderr)
            return 1

    if args.pc_members:
        try:
            members, fixed_preferences, _member_emails, chair_members, member_tags, member_tag_priority = load_pc_members(args.pc_members)
        except (FileNotFoundError, ValueError) as exc:
            print(exc, file=sys.stderr)
            return 1
        proposal_count = len(proposals)
        member_count = len(members)
        if args.max_first_per_member is not None and member_count * args.max_first_per_member < proposal_count:
            print(
                f"Insufficient primary-reviewer capacity: need {proposal_count} slots for {proposal_count} proposals, "
                f"but --max-first-per-member={args.max_first_per_member} with {member_count} members allows only "
                f"{member_count * args.max_first_per_member}. Increase the limit or add more members.",
                file=sys.stderr,
            )
            return 1
        if args.max_second_per_member is not None and member_count * args.max_second_per_member < proposal_count:
            print(
                f"Insufficient secondary-reviewer capacity: need {proposal_count} slots for {proposal_count} proposals, "
                f"but --max-second-per-member={args.max_second_per_member} with {member_count} members allows only "
                f"{member_count * args.max_second_per_member}. Increase the limit or add more members.",
                file=sys.stderr,
            )
            return 1
        if args.max_per_member is not None:
            total_required = proposal_count * args.reviewers_per_proposal
            total_capacity = member_count * args.max_per_member
            if total_capacity < total_required:
                print(
                    f"Insufficient reviewer capacity: assignments require {total_required} slots "
                    f"({proposal_count} proposals × {args.reviewers_per_proposal} reviewers), "
                    f"but --max-per-member={args.max_per_member} with {member_count} members allows only {total_capacity}.",
                    file=sys.stderr,
                )
                return 1
        # Process proposals with the fewest matching-tag members first so specialised
        # reviewers (e.g. the sole Pulsar expert) are not consumed by earlier proposals.
        if args.prefer_matching_tags and member_tags:
            original_order = {p["exp"]: i for i, p in enumerate(proposals)}

            def _match_count(proposal: Dict[str, Any]) -> tuple:
                ptags = set(proposal.get("science_tags") or [])
                if not ptags:
                    return (member_count, original_order[proposal["exp"]])
                n = sum(1 for tags in member_tags.values() if tags & ptags)
                return (n if n > 0 else member_count, original_order[proposal["exp"]])

            proposals.sort(key=_match_count)

        try:
            role_labels = generate_role_labels(args.reviewers_per_proposal)
            log_verbose(
                f"Assigning reviewers for {proposal_count} proposals using {member_count} PC members."
            )
            member_assignments = assign_reviewers(
                proposals,
                members,
                args.reviewers_per_proposal,
                args.max_per_member,
                fixed_preferences,
                args.max_first_per_member,
                args.max_second_per_member,
                chair_members,
                manual_conflicts or None,
                args.prefer_matching_tags,
                member_tags,
                member_tag_priority,
            )
        except ValueError as exc:
            print(exc, file=sys.stderr)
            return 1

        assignments_path = args.assignments or Path("reviewer_assignments.txt")
        try:
            log_verbose(f"Writing reviewer assignments to {assignments_path}.")
            write_assignments(proposals, assignments_path, role_labels or [])
        except OSError as exc:
            print(f"Failed to write assignments: {exc}", file=sys.stderr)
            return 1
        if args.member_summary and member_assignments is not None:
            try:
                log_verbose(f"Writing member summary to {args.member_summary}.")
                write_member_summary(member_assignments, args.member_summary, role_labels or [], member_tags)
            except OSError as exc:
                print(f"Failed to write member summary: {exc}", file=sys.stderr)
                return 1

    output_lines: List[str] = []

    for proposal in proposals:
        proposal.pop("normalised_text", None)
        proposal.pop("participants", None)
        additional_reviewers = proposal.get("additional_reviewers")
        rendered = list(
            render_record(
                proposal["exp"],
                proposal["pi"].title(),
                proposal["nets"],
                proposal["lambda"],
                proposal["title"],
                proposal.get("first_reviewer"),
                proposal.get("second_reviewer"),
                additional_reviewers,
            )
        )
        output_lines.extend(rendered)
        output_lines.append("")

    if args.output:
        log_verbose(f"Writing text output to {args.output}.")
        args.output.parent.mkdir(parents=True, exist_ok=True)
        content = "\n".join(output_lines)
        if output_lines and not content.endswith("\n"):
            content += "\n"
        args.output.write_text(content, encoding="utf-8")
    else:
        for line in output_lines:
            print(line)

    if args.agenda_docx:
        try:
            log_verbose(f"Writing agenda DOCX to {args.agenda_docx}.")
            write_docx_records([{"lines": build_agenda_lines()}], args.agenda_docx)
        except OSError as exc:
            print(f"Failed to write agenda DOCX: {exc}", file=sys.stderr)
            return 1

    if args.science_tags_file:
        try:
            log_verbose(f"Writing science tags to {args.science_tags_file}.")
            write_science_tags(proposals, args.science_tags_file)
        except OSError as exc:
            print(f"Failed to write science tags file: {exc}", file=sys.stderr)
            return 1

    if args.pi_emails_file:
        try:
            log_verbose(f"Writing PI emails to {args.pi_emails_file}.")
            write_pi_emails(proposals, args.pi_emails_file)
        except OSError as exc:
            print(f"Failed to write PI emails file: {exc}", file=sys.stderr)
            return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
