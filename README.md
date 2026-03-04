# EVN PC scripts

Utility scripts to help the EVN Programme Committee chair prepare reviewer templates, collect feedback, and summarise results.

## Prerequisites
- Python 3.9 or newer.
- Poppler's `pdftotext` binary available on `PATH` (used by `proposal_to_review_template.py`).
- Proposal PDFs exported from the EVN submission system.
- Optional: a current list of PC members in `EVN_pc_members.txt` (`Name Email` per line, plus optional `E25A001#1` fixed preferences).

## Quick start — `run-PC-scripts.sh`

The shell script `run-PC-scripts.sh` is the primary interface for the full PC-chair workflow. It wraps all Python scripts with the correct arguments for a given session. Before running, set the required environment variables:

| Variable | Purpose |
|---|---|
| `EVNPC_SHEETS` | Google Sheets CSV export URL (or local CSV path) for the review submission form |
| `GMAIL_ADDRESS` | Gmail address used as the SMTP sender for review reminders |
| `GMAIL_APPPWD` | Gmail App Password for SMTP authentication |

Edit the `SESSION` variable at the top of the script to match the current observing period (e.g. `2026A`).

### Commands

```
./run-PC-scripts.sh <section> [command] [--send]
```

| Section | Command | Description |
|---|---|---|
| `allocations` | `[session]` | Generate proposal-to-reviewer allocations. Prompts for session ID if omitted. |
| `primary` | `rename` | Rename primary review files from the Google Sheets CSV. |
| `primary` | `reminder` | Preview reminder emails for outstanding primary reviews (add `--send` to actually send). |
| `primary` | `latex` | Generate a LaTeX summary of primary reviews. |
| `secondary` | `rename` | Rename secondary review files from the Google Sheets CSV. |
| `secondary` | `reminder` | Preview reminder emails for outstanding secondary reviews (add `--send` to actually send). |
| `all` | `rename` | Rename all review files from the CSV. |
| `all` | `reminder` | Preview reminder emails for all outstanding reviews (add `--send` to actually send). |
| `all` | `latex` | Generate the combined LaTeX review summary. |
| `feedback` | | Generate draft feedback `.docx`. Add `--split-tex` to also write per-proposal LaTeX files. |

## Typical workflow

### 1. Generate reviewer allocations

Gather the proposal PDFs in a directory (e.g. `assessment/`) and an agenda `.docx` (`EVNPC_<session>_agenda.docx`).

Via the shell script:
```bash
./run-PC-scripts.sh allocations 2026A
```

Or directly:
```bash
python proposal_to_review_template.py \
  -p assessment -m EVN_pc_members.txt \
  -o EVNPC_2026A_assessment.txt \
  --reviewers-per-proposal 7 \
  --max-first-per-member 3 --max-second-per-member 3 --max-per-member 15 \
  --member-summary assignment_summary.html \
  --science-tags-file science-tags.txt --prefer-matching-tags \
  -A EVNPC_2026A_agenda.docx -v
```

This produces:
- `EVNPC_2026A_assessment.txt` — blank review template for each proposal
- `reviewer_assignments.txt` — CSV of reviewer allocations
- `assignment_summary.html` — per-reviewer summary table (paste into Outlook)
- `science-tags.txt` — inferred science categories per proposal

Key options:
- `--reviewers-per-proposal N` — number of reviewers per proposal (default: 2)
- `--max-per-member N`, `--max-first-per-member N`, `--max-second-per-member N` — load-balancing caps
- `--conflicts-file extra_conflicts.txt` — preload manual conflict exclusions (`E26A004: Alice Smith, Bob Jones`)
- `--prefer-matching-tags` — prefer reviewers whose declared expertise matches the proposal science tags
- Add `*` to a PC member's surname in `EVN_pc_members.txt` to mark them as a fallback chair (e.g. `Jack Radcliffe*`)

### 2. Distribute and collect reviews

Share the generated template with PC members. After receiving Google Form responses, download the CSV and uploaded files, then rename them:

```bash
./run-PC-scripts.sh primary rename
./run-PC-scripts.sh secondary rename
# or for all at once:
./run-PC-scripts.sh all rename
```

Or directly:
```bash
python rename_reviews_from_csv.py "$EVNPC_SHEETS" \
  --prefix primary \
  --source-dir "Copy of EVN.../Review submission (File responses)" \
  --dest-dir primary_pc_reviews --prefer-newest
```

Use `--dry-run` to preview changes and `--skip-missing` to ignore missing files.

### 3. Send review reminders

Before sending reminders, create a `due_dates.txt` file in the working directory:

```
First review: 2026-02-15
Second review: 2026-02-28
Additional: 2026-02-28
```

Preview reminders (dry-run, no emails sent):
```bash
./run-PC-scripts.sh primary reminder
```

Send reminders:
```bash
./run-PC-scripts.sh primary reminder --send
```

Or directly:
```bash
python review_reminder.py \
  --assignments reviewer_assignments.txt \
  --reviews-dir primary_pc_reviews \
  --pc-members EVN_pc_members.txt \
  --due-dates due_dates.txt \
  --smtp-username $GMAIL_ADDRESS --smtp-password $GMAIL_APPPWD \
  [--dry-run]
```

Save reminder drafts to files instead of sending:
```bash
python review_reminder.py ... --export-emails drafts_dir/
```

### 4. Build the LaTeX review summary

```bash
./run-PC-scripts.sh all latex
```

Or directly:
```bash
python reviews_to_latex.py \
  -r all_pc_reviews \
  -o EVNPC_2026A_review_summary.tex \
  -a reviewer_assignments.txt \
  --agenda-txt EVNPC_agenda_items_2026A.txt \
  --code-mapping evn_code_mapping.txt \
  -t "EVN PC 2026A review summary"
```

The output LaTeX file lists each proposal with reviewer initials, grades, and comments. Supply `--title` and `--version` to customize the document header.

### 5. Generate feedback emails

After the PC meeting, fill in the consensus grades and comments in the assessment file, then generate draft feedback emails:

```bash
./run-PC-scripts.sh feedback
# or with per-proposal LaTeX files:
./run-PC-scripts.sh feedback --split-tex
```

Or directly:
```bash
python generate_feedback_emails.py \
  -a EVNPC_2026A_assessment.txt \
  -m evn_code_mapping.txt \
  -r reviewer_assignments.txt \
  -o EVNPC_2026A_feedback.docx \
  --suffix-file evn_pc_suffix_content.txt \
  --session 2026A \
  [--split-tex feedback_tex/ --reviews-dir all_pc_reviews/]
```

This produces a `.docx` with one draft email per page (including a TOC), with primary/secondary reviewer names annotated as Word comments. With `--split-tex`, it also writes per-proposal standalone LaTeX files to the specified directory.

## Additional notes

- `template.py` contains the low-level renderer used by `proposal_to_review_template.py`; it can also be used directly by piping ampersand-delimited records from legacy sources.
- Sample member data is provided in `EVN_pc_members.txt`, including example email addresses. Update it each cycle so the automatic assignment logic remains accurate.
- Run any script with `-h`/`--help` to view the full set of options and usage examples.
