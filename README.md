# EVN PC scripts

Utility scripts to help the EVN Programme Committee chair prepare reviewer templates, collect feedback, and summarise results.

## Prerequisites
- Python 3.9 or newer.
- Poppler's `pdftotext` binary available on `PATH` (used by `proposal_to_review_template.py`).
- Proposal PDFs exported from the EVN submission system.
- Optional: a current list of PC members in `EVN_pc_members.txt` — see [PC members file format](#pc-members-file-format) below.

### Additional packages for `supplement_meeting_notes.py`
```
pip install python-docx mlx-whisper
```
- `mlx-whisper` is only needed if transcribing audio files (Apple Silicon).
- AI synthesis requires the `claude` CLI (Claude Code) to be installed and on `PATH`.

## Quick start — `run-PC-scripts.sh`

The shell script `run-PC-scripts.sh` is the primary interface for the full PC-chair workflow. It wraps all Python scripts with the correct arguments for a given session. Before running, set the required environment variables:

| Variable | Purpose |
|---|---|
| `EVNPC_SHEETS` | Google Sheets CSV export URL (or local CSV path) for the review submission form |
| `GMAIL_ADDRESS` | Gmail address used as the SMTP sender for review reminders and feedback emails |
| `GMAIL_APP_PWD` | Gmail App Password for SMTP/IMAP authentication |

Edit the `SESSION` variable at the top of the script to match the current observing period (e.g. `2026A`). The script also assumes the scripts live in `SCRIPTS_DIR` and reviewer file uploads land in `REVIEWS_SOURCE_DIR` (both set near the top of `run-PC-scripts.sh`) and runs everything through the pinned `PY_EXEC` interpreter — adjust these paths for your machine.

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
| `secondary` | `latex` | Generate a LaTeX summary of secondary reviews. |
| `all` | `rename` | Rename all review files from the CSV. |
| `all` | `reminder` | Preview reminder emails for all outstanding reviews (add `--send` to actually send). |
| `all` | `latex` | Generate the combined LaTeX review summary. |
| `feedback` | | Generate draft feedback `.docx`. Add `--split-tex` to also write per-proposal LaTeX files. |
| `feedback` | `--tex-only` | Write per-proposal LaTeX files only — no `.docx` is produced. |
| `feedback` | `reminder` | Preview reminder emails for PC members with missing feedback summaries (add `--send` to actually send). |
| `feedback` | `send` | Preview feedback emails to PIs (dry-run). Add `--draft` to save as Gmail drafts, or `--send` to deliver. |

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
- `EVNPC_2026A_pi_emails.txt` — `CODE: email` file used later by `send_feedback_emails.py`

Key options:
- `--reviewers-per-proposal N` — number of reviewers per proposal (default: 2)
- `--max-per-member N`, `--max-first-per-member N`, `--max-second-per-member N` — load-balancing caps
- `--conflicts-file extra_conflicts.txt` — preload manual conflict exclusions (`E26A004: Alice Smith, Bob Jones`)
- `--prefer-matching-tags` — prefer reviewers whose declared expertise matches the proposal science tags (see [PC members file format](#pc-members-file-format))
- Add `*` to any part of a PC member's name in `EVN_pc_members.txt` to mark them as a fallback chair who receives leftover assignments (e.g. `Jack Radcliffe*`)

#### PC members file format

`EVN_pc_members.txt` has one member per line. Each line is whitespace-separated and may combine any of the following tokens in any order:

```
Jack Radcliffe* jack.f.radcliffe@gmail.com E25A001#1 E25A016#2 AGN astrometry
John Smith john.smith@example.org galactic maser
Sarah Marais
```

| Token | Meaning |
|---|---|
| Name words | The member's display name. Append `*` to any word to mark a fallback chair. |
| A token containing `@` | The member's email address (used for reminders). Optional. |
| `CODE#1` / `CODE#2` | Fixed assignment: force this member to be the primary (`#1`) or secondary (`#2`) reviewer of proposal `CODE`. May be repeated. |
| Science keywords | Declared expertise tags, used by `--prefer-matching-tags`. Listed left-to-right in priority order (first = highest priority). |

Recognised science keywords map to these categories: **Galactic** (`galactic`, `milky way`, `stellar`, `star formation`, `pulsar`), **Extragalactic** (`extragalactic`, `galaxy`, `galaxies`, `cluster`), **Spectral Line** (`spectral line`, `molecular line`, `hi`), **Maser** (`maser`, `megamaser`, `ohm`, `water maser`, `methanol maser`), **Transient** (`transient`, `burst`, `frb`, `grb`), **AGN** (`agn`, `blazar`, `seyfert`, `quasar`), **Supernovae** (`supernova`, `snr`, `remnant`), **Pulsar** (`pulsar`, `psr`, `scintillation`, `magnetar`, `neutron star`, `msp`), **Astrometry** (`astrometry`, `proper motion`, `parallax`, `reference frame`, `icrf`, `geodesy`), and **Other**. Proposal science tags are inferred automatically from each PDF (written to `--science-tags-file`); with `--prefer-matching-tags` the allocator preferentially matches proposals to members whose declared tags overlap.

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

Key options:
- `--dry-run` — preview renames without touching files
- `--skip-missing` — ignore CSV rows whose uploaded file is missing
- `--prefer-newest` — when a reviewer submits more than once, keep the most recent file
- `--include-unlisted` — also pick up `... - First Last.ext` files in the source directory that have no matching CSV row (e.g. reviews emailed in and dropped into the folder manually). Used by `run-PC-scripts.sh`.
- `--no-docx-conversion` — disable the automatic conversion of `.docx` submissions to `.txt`

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

Other useful options:
- `--summary-only` — print only the outstanding-reviews summary report, with no per-reviewer reminder content
- `--current-date 2026-02-20` — override "now" (ISO format) when computing overdue status, for testing or backfilling
- `--message-template` / `--subject-template` — customise the reminder body and subject
- `--smtp-server`, `--smtp-port`, `--smtp-use-ssl`, `--from-address`, `--cc`, `--reply-to` — SMTP and header overrides

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
# also write per-proposal LaTeX files:
./run-PC-scripts.sh feedback --split-tex
# write LaTeX files only (skip the .docx entirely):
./run-PC-scripts.sh feedback --tex-only
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
  [--split-tex feedback_tex/ --reviews-dir all_pc_reviews/] \
  [--tex-only]
```

This produces a `.docx` with one draft email per page (including a TOC), with primary/secondary reviewer names annotated as Word comments. With `--split-tex`, it also writes per-proposal standalone LaTeX files to the specified directory. Use `--tex-only` (together with `--split-tex`) to skip the `.docx` entirely and only produce the LaTeX files.

### 5b. Send feedback reminders

After distributing draft feedback emails, send reminders to PC members who have not yet returned their feedback summaries:

```bash
./run-PC-scripts.sh feedback reminder
# add --send to actually send:
./run-PC-scripts.sh feedback reminder --send
```

Or directly:
```bash
python review_reminder.py \
  --feedback-docx EVNPC_2026A_feedback.docx \
  --assignments reviewer_assignments.txt \
  --pc-members EVN_pc_members.txt \
  --smtp-username $GMAIL_ADDRESS --smtp-password $GMAIL_APPPWD \
  [--dry-run]
```

### 5c. Send feedback emails to PIs

Once the feedback `.docx` has been reviewed and all placeholders filled in, compile the per-proposal PDFs and send the emails to PIs.

Preview (dry-run, no emails sent):
```bash
./run-PC-scripts.sh feedback send
```

Save as Gmail drafts (review in your Drafts folder before sending):
```bash
./run-PC-scripts.sh feedback send --draft
```

Send:
```bash
./run-PC-scripts.sh feedback send --send
```

Or directly:
```bash
python send_feedback_emails.py \
  --feedback-docx EVNPC_2026A_feedback.docx \
  --pi-emails-file EVNPC_2026A_pi_emails.txt \
  --pdf-dir feedback_tex/ \
  --code-mapping evn_code_mapping.txt \
  --smtp-username $GMAIL_ADDRESS --smtp-password $GMAIL_APPPWD \
  [--dry-run | --draft | --send]
```

The script automatically compiles any stale `.tex` files in `--pdf-dir` using `pdflatex` before sending. Each email includes a plain-text and HTML body (preserving bold/italic from the `.docx`) and the compiled PDF as an attachment. Proposals with unresolved choice placeholders (e.g. `[was / was not]`) are skipped automatically.

Key options:
- `--proposals CODE ...` — send only for specific proposal codes (useful for testing or resending individual emails)
- `--skip-compile` — skip LaTeX compilation if PDFs are already up-to-date
- `--sent-log FILE` — JSON log tracking which emails have been delivered; proposals already in the log are skipped on re-runs
- `--force` — resend even if a proposal is already recorded in the sent log
- `--draft` — save to Gmail Drafts via IMAP instead of sending
- `--cc ADDRESS` — add CC recipients (may be repeated)
- `--reply-to ADDRESS` — set a Reply-To header
- `--export-emails DIR` — save each email as a `.eml` file

### 6. Supplement meeting notes

After the PC meeting, supplement the agenda/notes DOCX with content from Zoom closed captions, the chat log, and/or audio recordings. Claude synthesises the sources and inserts contextual notes inline:

```bash
python supplement_meeting_notes.py EVNPC_2026A_agenda.docx \
  --captions closed_caption.txt \
  --chat chat.txt \
  --audio audio1.m4a audio2.m4a
```

This produces `EVNPC_2026A_agenda_supplemented.docx` with colour-coded inline notes (captions = blue, chat = purple, audio = green).

Key options:
- `--output FILE` — custom output path
- `--model MODEL` — Whisper model for transcription (default: `mlx-community/whisper-large-v3-turbo`)
- `--no-ai` — skip Claude synthesis; just append full transcripts as a new section

Audio transcripts are cached as `<audio>.transcript.txt` alongside the source file, so re-runs do not re-transcribe.

## Additional notes

- `template.py` contains the low-level renderer used by `proposal_to_review_template.py`; it can also be used directly by piping ampersand-delimited records from legacy sources.
- Sample member data is provided in `EVN_pc_members.txt`, including example email addresses. Update it each cycle so the automatic assignment logic remains accurate.
- Run any script with `-h`/`--help` to view the full set of options and usage examples.
