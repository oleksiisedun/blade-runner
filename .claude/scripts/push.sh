#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(git -C "$(dirname "$0")" rev-parse --show-toplevel)"

ENV_FILE="$REPO_ROOT/.env"
if [[ ! -f "$ENV_FILE" ]]; then
  echo "Error: .env file not found at $ENV_FILE" >&2
  exit 1
fi
# shellcheck source=/dev/null
source "$ENV_FILE"

GMAIL_USER="oleksiisedun@gmail.com"
RECIPIENTS=("oleksiisedun@gmail.com")

send_notification() {
  local commit_hash="$1"
  local commit_msg="$2"
  local pushed_at="$3"

  for rcpt in "${RECIPIENTS[@]}"; do
    curl --silent --ssl-reqd \
      --url "smtps://smtp.gmail.com:465" \
      --user "${GMAIL_USER}:${GMAIL_PASS}" \
      --mail-from "${GMAIL_USER}" \
      --mail-rcpt "${rcpt}" \
      --upload-file - <<EOF
From: Apps Script Deploy <${GMAIL_USER}>
To: ${rcpt}
Subject: Deployed: ${commit_msg}
Content-Type: text/plain; charset=utf-8

Project was pushed to Google Apps Script and git.

Commit : ${commit_hash}
Message: ${commit_msg}
Time   : ${pushed_at}
EOF
  done
}

echo "==> Pushing to Google Apps Script via clasp..."
(cd "$REPO_ROOT" && clasp push)

echo ""
echo "==> Pushing to git..."
cd "$REPO_ROOT"

if git diff --quiet && git diff --cached --quiet; then
  echo "Nothing to commit, working tree clean."
else
  git add .
  git commit -m "Update script"
fi

git push

COMMIT_HASH="$(git rev-parse --short HEAD)"
COMMIT_MSG="$(git log -1 --pretty=%s)"
PUSHED_AT="$(date -u '+%Y-%m-%d %H:%M UTC')"

echo ""
printf "==> Send deployment notification? [y/N] "
read -r NOTIFY
if [[ "$NOTIFY" =~ ^[Yy]$ ]]; then
  echo "Sending..."
  send_notification "$COMMIT_HASH" "$COMMIT_MSG" "$PUSHED_AT"
  echo "Notified: ${RECIPIENTS[*]}"
else
  echo "Skipped."
fi

echo ""
echo "Done."
