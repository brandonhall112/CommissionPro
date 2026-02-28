#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$repo_root"

remote_url="$(git config --get remote.origin.url || true)"
if [[ -z "$remote_url" ]]; then
  echo "origin remote is not configured. Run ./scripts/ensure-origin-remote.sh first." >&2
  exit 1
fi

if [[ "$remote_url" =~ github.com[:/]([^/]+)/([^/.]+)(\.git)?$ ]]; then
  owner="${BASH_REMATCH[1]}"
  repo="${BASH_REMATCH[2]}"
else
  echo "origin is not a GitHub URL: $remote_url" >&2
  exit 1
fi

branch="$(git rev-parse --abbrev-ref HEAD)"
base="${1:-main}"

url="https://github.com/${owner}/${repo}/compare/${base}...${branch}?expand=1"
echo "$url"
