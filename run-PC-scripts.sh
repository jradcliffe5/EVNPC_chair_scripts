#!/bin/bash

usage() {
    echo "Usage: ./code_run.sh <section> [command]"
    echo ""
    echo "Sections and commands:"
    echo "  allocations [session]    -- run proposal-to-review allocation (prompts for session if omitted)"
    echo "  primary   rename         -- rename primary reviews from CSV"
    echo "  primary   reminder [--send]  -- send primary review reminders (dry-run by default)"
    echo "  primary   latex              -- generate primary review LaTeX summary"
    echo "  secondary rename             -- rename secondary reviews from CSV"
    echo "  secondary reminder [--send]  -- send secondary review reminders (dry-run by default)"
    echo "  secondary latex              -- generate secondary review LaTeX summary"
    echo "  all       rename             -- rename all reviews from CSV"
    echo "  all       reminder [--send]  -- send all review reminders (dry-run by default)"
    echo "  all       latex          -- generate all-review LaTeX summary"
    echo "  feedback  [--split-tex]  -- generate draft feedback emails .docx (optionally split per-proposal .tex)"
  echo "  feedback  --tex-only     -- generate per-proposal .tex files only (no .docx written)"
  echo "  feedback  reminder [--send]  -- send reminders to PC members with missing feedback summaries"
}

SESSION="2026A"

SCRIPTS_DIR="../../pc_chair/EVNPC_chair_scripts"                                                                                                                                                                                                                
SHEETS_URL=$EVNPC_SHEETS
REVIEWS_SOURCE_DIR="../../pc_chair/Copy of EVN programme committee review submission (File responses)/Review submission (File responses)"   
PY_EXEC="$HOME/.pyenv/versions/3.13.7/bin/python"

SECTION=$1
CMD=$2
FLAG=$3

# --send removes --dry-run from reminder commands
DRY_RUN="--dry-run"
if [ "$FLAG" = "--send" ]; then
  DRY_RUN=""
fi

case "$SECTION" in

  allocations)
    SESSION=$CMD
    if [ -z "$SESSION" ]; then
      read -rp "Enter session ID (e.g. 2026A): " SESSION
    fi
    $PY_EXEC ${SCRIPTS_DIR}/proposal_to_review_template.py \
      -m EVN_pc_members.txt -p assessment -o "EVNPC_${SESSION}_assessment.txt" \
      --reviewers-per-proposal 7 --max-first-per-member 3 --max-second-per-member 3 \
      --max-per-member 15 --member-summary assignment_summary.html \
      --science-tags-file science-tags.txt --prefer-matching-tags \
      -A "EVNPC_${SESSION}_agenda.docx" -v
    ;;

  primary)
    case "$CMD" in
      rename)
        $PY_EXEC ${SCRIPTS_DIR}/rename_reviews_from_csv.py \
          "$SHEETS_URL" \
          --prefix primary \
          --source-dir "$REVIEWS_SOURCE_DIR" \
          --dest-dir primary_pc_reviews --prefer-newest
        ;;
      reminder)
        $PY_EXEC ${SCRIPTS_DIR}/review_reminder.py \
          --assignments reviewer_assignments.txt --reviews-dir primary_pc_reviews \
          --pc-members EVN_pc_members.txt --due-dates due_dates.txt \
          --smtp-username $GMAIL_ADDRESS --smtp-password $GMAIL_APPPWD $DRY_RUN
        ;;
      latex)
        $PY_EXEC ${SCRIPTS_DIR}/reviews_to_latex.py \
          -r primary_pc_reviews -o review_primary_summary.tex -a reviewer_assignments.txt
        ;;
      *)
        echo "Unknown command '$CMD' for section 'primary'."
        usage; exit 1
        ;;
    esac
    ;;

  secondary)
    case "$CMD" in
      rename)
        $PY_EXEC ${SCRIPTS_DIR}/rename_reviews_from_csv.py \
          "$SHEETS_URL" \
          --prefix secondary \
          --source-dir "$REVIEWS_SOURCE_DIR" \
          --dest-dir secondary_pc_reviews --prefer-newest
        ;;
      reminder)
        $PY_EXEC ${SCRIPTS_DIR}/review_reminder.py \
          --assignments reviewer_assignments.txt --reviews-dir secondary_pc_reviews \
          --pc-members EVN_pc_members.txt --due-dates due_dates.txt \
          --smtp-username $GMAIL_ADDRESS --smtp-password $GMAIL_APPPWD $DRY_RUN
        ;;
      latex)
        $PY_EXEC ${SCRIPTS_DIR}/reviews_to_latex.py \
          -r secondary_pc_reviews -o review_secondary_summary.tex -a reviewer_assignments.txt
        ;;
      *)
        echo "Unknown command '$CMD' for section 'secondary'."
        usage; exit 1
        ;;
    esac
    ;;

  all)
    case "$CMD" in
      rename)
        $PY_EXEC ${SCRIPTS_DIR}/rename_reviews_from_csv.py \
          "$SHEETS_URL" \
          --prefix all \
          --source-dir "$REVIEWS_SOURCE_DIR" \
          --dest-dir all_pc_reviews --prefer-newest
        ;;
      reminder)
        $PY_EXEC ${SCRIPTS_DIR}/review_reminder.py \
          --assignments reviewer_assignments.txt --reviews-dir all_pc_reviews \
          --pc-members EVN_pc_members.txt --due-dates due_dates.txt \
          --smtp-username $GMAIL_ADDRESS --smtp-password $GMAIL_APPPWD $DRY_RUN
        ;;
      latex)
        $PY_EXEC ${SCRIPTS_DIR}/reviews_to_latex.py \
          -r all_pc_reviews -o EVNPC_review_summary_$SESSION".tex" -a reviewer_assignments.txt \
          --agenda-txt EVNPC_agenda_items_$SESSION".txt" \
          --code-mapping evn_code_mapping.txt \
          -t "EVN PC ${SESSION} review summary"
        ;;
      *)
        echo "Unknown command '$CMD' for section 'all'."
        usage; exit 1
        ;;
    esac
    ;;

  feedback)
    case "$CMD" in
      reminder)
        $PY_EXEC ${SCRIPTS_DIR}/review_reminder.py \
          --feedback-docx "EVNPC_${SESSION}_feedback.docx" \
          --assignments reviewer_assignments.txt \
          --pc-members EVN_pc_members.txt \
          --smtp-username $GMAIL_ADDRESS --smtp-password $GMAIL_APPPWD $DRY_RUN
        ;;
      --tex-only)
        $PY_EXEC ${SCRIPTS_DIR}/generate_feedback_emails.py \
          -a "EVNPC_${SESSION}_assessment.txt" \
          -m evn_code_mapping.txt \
          -r reviewer_assignments.txt \
          -o "EVNPC_${SESSION}_feedback.docx" \
          --suffix-file evn_pc_suffix_content.txt \
          --session "${SESSION}" \
          --split-tex feedback_tex/ --reviews-dir all_pc_reviews \
          --tex-only
        ;;
      *)
        read -rp "This will overwrite EVNPC_${SESSION}_feedback.docx. Type 'yes' to confirm: " CONFIRM
        if [ "$CONFIRM" != "yes" ]; then
          echo "Aborted."; exit 1
        fi
        SPLIT_TEX_ARGS=""
        if [ "$CMD" = "--split-tex" ]; then
          SPLIT_TEX_ARGS="--split-tex feedback_tex/ --reviews-dir all_pc_reviews"
        fi
        $PY_EXEC ${SCRIPTS_DIR}/generate_feedback_emails.py \
          -a "EVNPC_${SESSION}_assessment.txt" \
          -m evn_code_mapping.txt \
          -r reviewer_assignments.txt \
          -o "EVNPC_${SESSION}_feedback.docx" \
          --suffix-file evn_pc_suffix_content.txt \
          --session "${SESSION}" \
          $SPLIT_TEX_ARGS
        ;;
    esac
    ;;

  *)
    echo "Unknown section '$SECTION'."
    usage; exit 1
    ;;

esac
