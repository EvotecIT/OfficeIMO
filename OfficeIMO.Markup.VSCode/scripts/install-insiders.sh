#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$repo_root"

if ! command -v pwsh >/dev/null 2>&1; then
  echo "pwsh is required to build the VSIX with scripts/package-vsix.ps1" >&2
  exit 1
fi

pwsh -NoProfile -File "$repo_root/scripts/package-vsix.ps1" -OutputDirectory "$repo_root"

vsix_name="$(node -e "const p=require('./package.json'); console.log(p.name + '-' + p.version + '.vsix')")"
code-insiders --install-extension "$repo_root/$vsix_name" "${1:-}"
echo "Installed. Reload VS Code Insiders to activate the update."
