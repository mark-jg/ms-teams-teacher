#!/usr/bin/env bash
# Nieuwe versie uitbrengen.
#   1. Zet een regel "- X.Y.Z: ..." bovenaan "## Wijzigingen" in README.md en commit je werk.
#   2. scripts/release.sh X.Y.Z
# Het script zet het versienummer in manifest.json en README.md, commit, tagt vX.Y.Z en pusht.
# GitHub Actions bouwt daarna de zip en de release; de geïnstalleerde extensies zien de update binnen 6 uur.
set -euo pipefail
cd "$(dirname "$0")/.."

V="${1:?Gebruik: scripts/release.sh X.Y.Z}"
[[ "$V" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]] || { echo "Ongeldige versie: $V (verwacht X.Y.Z)"; exit 1; }
[ -z "$(git status --porcelain)" ] || { echo "Er zijn niet-gecommitte wijzigingen. Commit die eerst."; exit 1; }
git rev-parse -q --verify "refs/tags/v$V" >/dev/null && { echo "Tag v$V bestaat al."; exit 1; }
grep -qE "^- ${V//./\\.}:" README.md || { echo "Voeg eerst een regel '- $V: ...' toe onder '## Wijzigingen' in README.md."; exit 1; }

sed -i '' -E "s/\"version\": \"[0-9.]+\"/\"version\": \"$V\"/" manifest.json
sed -i '' -E "1s/\(v[0-9.]+\)/(v$V)/" README.md

git add manifest.json README.md
git commit -q -m "Versie $V"
git tag -a "v$V" -m "Versie $V"
git push -q origin main "v$V"
echo "v$V gepusht. GitHub Actions bouwt de release: https://github.com/mark-jg/ms-teams-teacher/actions"
