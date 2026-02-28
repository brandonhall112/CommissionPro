#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$repo_root"

if git config --get remote.origin.url >/dev/null 2>&1; then
  echo "origin already configured: $(git config --get remote.origin.url)"
  exit 0
fi

repo_name="$(basename "$repo_root")"
owner=""

# Infer owner from merge commit subjects like:
#   Merge pull request #57 from brandonhall112/codex/...
while IFS= read -r subject; do
  if [[ "$subject" =~ from[[:space:]]+([^/]+)/ ]]; then
    owner="${BASH_REMATCH[1]}"
    break
  fi
done < <(git log --merges --pretty=%s -n 100)

if [[ -z "$owner" ]]; then
  echo "Could not infer GitHub owner from merge history; leaving remote unconfigured."
  exit 1
fi

origin_url="https://github.com/${owner}/${repo_name}.git"
git remote add origin "$origin_url"

echo "Configured origin: $origin_url"
