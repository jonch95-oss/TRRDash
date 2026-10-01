#!/usr/bin/env bash
# Double-click this file in Finder to refresh the dashboard.
#
# It finds the newest TRR inventory export in ~/Downloads, shows you the
# numbers before changing anything, and only then rebuilds and publishes.
# Nothing leaves your machine except the finished index.html.

set -uo pipefail
cd "$(dirname "$0")" || exit 1

DOWNLOADS="${HOME}/Downloads"
PATTERN="*Inventory*Export*.xlsx"

say()  { printf '\n\033[1m%s\033[0m\n' "$*"; }
fail() { printf '\n\033[31m%s\033[0m\n' "$*"; printf '\nPress return to close.\n'; read -r _; exit 1; }

# --- interpreter: prefer the project venv, fall back to system python3 ---
if   [ -x ".venv/bin/python" ]; then PY=".venv/bin/python"
elif [ -x "venv/bin/python" ];  then PY="venv/bin/python"
elif command -v python3 >/dev/null 2>&1; then PY="python3"
else fail "No Python found. Install Python 3, then run: pip install -r requirements.txt"
fi

"$PY" -c "import openpyxl" 2>/dev/null || fail \
  "openpyxl is missing. Run this once in Terminal, from this folder:
    $PY -m pip install -r requirements.txt"

# --- the wholesale invoice must already be in place ---
[ -f "data/sales.xlsx" ] || fail \
  "data/sales.xlsx is missing.
Copy your Icon wholesale invoice there once — it is the source of all brand
and cost data, and the dashboard cannot be built without it."

# --- newest matching export in Downloads ---
# Plain glob + -nt comparison: portable, space-safe, and (unlike `ls | head`)
# it cannot silently fall back to listing the current directory when nothing matches.
NEWEST=""
shopt -s nullglob
for f in "$DOWNLOADS"/$PATTERN; do
  [ -f "$f" ] || continue
  if [ -z "$NEWEST" ] || [ "$f" -nt "$NEWEST" ]; then NEWEST="$f"; fi
done
shopt -u nullglob
[ -n "$NEWEST" ] || fail \
  "No file matching '$PATTERN' found in $DOWNLOADS.
Download this week's export first, or copy it to data/inventory.xlsx yourself."

say "Using export:"
printf '   %s\n   modified %s\n' "$(basename "$NEWEST")" "$(ls -l "$NEWEST" | awk '{print $6, $7, $8}')"

# --- preview before touching anything ---
cp "$NEWEST" data/inventory.xlsx || fail "Could not copy the export into data/."
say "Checking the numbers (nothing has changed yet)…"
"$PY" refresh.py --dry-run || fail \
  "The check failed — see the message above. index.html was NOT modified."

say "Publish these numbers? [y/N]"
read -r REPLY
case "$REPLY" in
  [yY]*) ;;
  *) printf '\nStopped. index.html is untouched.\n\nPress return to close.\n'; read -r _; exit 0 ;;
esac

say "Rebuilding…"
"$PY" refresh.py || fail "The refresh failed — see above. index.html was NOT modified."

say "Publishing…"
git add index.html history.json schema.json 2>/dev/null
if git diff --cached --quiet; then
  printf 'Nothing changed since the last refresh — no commit needed.\n'
else
  git commit -q -m "refresh $(date +%F)" && git push -q origin HEAD \
    && printf 'Pushed. Vercel will redeploy in about 20 seconds.\n' \
    || fail "Commit or push failed — index.html is rebuilt, it just was not published.
If git complained about your identity, set it once in Terminal:
    git config --global user.name  \"Your Name\"
    git config --global user.email \"you@example.com\"
Then double-click this file again."
fi

printf '\n\033[32mDone.\033[0m https://trr-dashboard-jonch95-oss-projects.vercel.app\n'
printf '\nPress return to close.\n'; read -r _
