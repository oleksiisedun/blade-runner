#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(git -C "$(dirname "$0")" rev-parse --show-toplevel)"

GMAIL_USER="oleksiisedun@gmail.com"
GMAIL_PASS="lbwuqoqqcqfraxij"
RECIPIENTS=("oleksiisedun@gmail.com" "3blukas1@gmail.com")

# Resolve target directory
if [[ $# -ge 1 ]]; then
  TARGET_DIR="$REPO_ROOT/$1"
  SUBPROJECT="$1"
  if [[ ! -d "$TARGET_DIR" ]]; then
    echo "Error: subproject '$1' not found in repo root." >&2
    exit 1
  fi
elif [[ -f "$REPO_ROOT/.clasp.json" ]]; then
  TARGET_DIR="$REPO_ROOT"
  SUBPROJECT="."
else
  echo "No subproject specified and no .clasp.json found in repo root."
  echo "Available subprojects:"
  for d in "$REPO_ROOT"/*/; do
    [[ -f "$d/.clasp.json" ]] && echo "  $(basename "$d")"
  done
  printf "Enter subproject name: "
  read -r SUBPROJECT
  TARGET_DIR="$REPO_ROOT/$SUBPROJECT"
  if [[ ! -d "$TARGET_DIR" ]]; then
    echo "Error: subproject '$SUBPROJECT' not found." >&2
    exit 1
  fi
fi

send_notification() {
  local subproject="$1"
  local commit_hash="$2"
  local commit_msg="$3"
  local pushed_at="$4"

  for rcpt in "${RECIPIENTS[@]}"; do
    curl --silent --ssl-reqd \
      --url "smtps://smtp.gmail.com:465" \
      --user "${GMAIL_USER}:${GMAIL_PASS}" \
      --mail-from "${GMAIL_USER}" \
      --mail-rcpt "${rcpt}" \
      --upload-file - <<EOF
From: Apps Script Deploy <${GMAIL_USER}>
To: ${rcpt}
Subject: ${subproject} deployed: ${commit_msg}
Content-Type: text/plain; charset=utf-8

${subproject} was pushed to Google Apps Script and git.

Commit : ${commit_hash}
Message: ${commit_msg}
Time   : ${pushed_at}
EOF
  done
}

echo "==> Pushing '$SUBPROJECT' to Google Apps Script via clasp..."
(cd "$TARGET_DIR" && clasp push)

echo ""
echo "==> Pushing to git..."
cd "$REPO_ROOT"

if git diff --quiet && git diff --cached --quiet; then
  echo "Nothing to commit, working tree clean."
else
  if [[ "$SUBPROJECT" == "." ]]; then
    git add .
  else
    git add "$SUBPROJECT/"
  fi
  git commit -m "Update ${SUBPROJECT} script"
fi

git push

COMMIT_HASH="$(git rev-parse --short HEAD)"
COMMIT_MSG="$(git log -1 --pretty=%s)"
PUSHED_AT="$(date -u '+%Y-%m-%d %H:%M UTC')"

echo ""
echo "==> Sending deployment notifications..."
send_notification "$SUBPROJECT" "$COMMIT_HASH" "$COMMIT_MSG" "$PUSHED_AT"
echo "Notified: ${RECIPIENTS[*]}"

echo ""
echo "Done."
