#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(git -C "$(dirname "$0")" rev-parse --show-toplevel)"
MAILER_DIR="$REPO_ROOT/Mailer"

GMAIL_USER="oleksiisedun@gmail.com"
GMAIL_PASS="lbwuqoqqcqfraxij"
RECIPIENTS=("oleksiisedun@gmail.com" "kogut.alexandr83@gmail.com")

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
From: Mailer Deploy <${GMAIL_USER}>
To: ${rcpt}
Subject: Mailer deployed: ${commit_msg}
Content-Type: text/plain; charset=utf-8

Mailer was pushed to Google Apps Script and git.

Commit : ${commit_hash}
Message: ${commit_msg}
Time   : ${pushed_at}
EOF
  done
}

echo "==> Pushing to Google Apps Script via clasp..."
(cd "$MAILER_DIR" && clasp push)

echo ""
echo "==> Pushing to git..."
cd "$REPO_ROOT"

if git diff --quiet && git diff --cached --quiet; then
  echo "Nothing to commit, working tree clean."
else
  git add Mailer/
  git commit -m "Update Mailer script"
fi

git push

COMMIT_HASH="$(git rev-parse --short HEAD)"
COMMIT_MSG="$(git log -1 --pretty=%s)"
PUSHED_AT="$(date -u '+%Y-%m-%d %H:%M UTC')"

echo ""
echo "==> Sending deployment notifications..."
send_notification "$COMMIT_HASH" "$COMMIT_MSG" "$PUSHED_AT"
echo "Notified: ${RECIPIENTS[*]}"

echo ""
echo "Done."
