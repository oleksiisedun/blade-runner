#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(git -C "$(dirname "$0")" rev-parse --show-toplevel)"
MAILER_DIR="$REPO_ROOT/Mailer"

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
echo ""
echo "Done."
