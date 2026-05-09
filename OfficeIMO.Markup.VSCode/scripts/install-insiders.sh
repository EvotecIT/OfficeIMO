#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$repo_root"

npm run package:vsix -- --output-directory "$repo_root"

vsix_name="$(node -e "const p=require('./package.json'); console.log(p.name + '-' + p.version + '.vsix')")"
code-insiders --install-extension "$repo_root/$vsix_name" "${1:-}"
echo "Installed. Reload VS Code Insiders to activate the update."
