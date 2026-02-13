#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$repo_root"

git config core.hooksPath .githooks
echo "Configured core.hooksPath=.githooks"
echo "Active pre-commit hook: $repo_root/.githooks/pre-commit"
