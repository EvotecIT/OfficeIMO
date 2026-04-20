#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
stable=0
force=0

for arg in "$@"; do
  case "$arg" in
    --stable) stable=1 ;;
    --force) force=1 ;;
  esac
done

name="$(node -e "const p=require('$repo_root/package.json'); console.log(p.publisher + '.' + p.name + '-' + p.version)")"
ext_id="$(node -e "const p=require('$repo_root/package.json'); console.log(p.publisher + '.' + p.name)")"

if [[ "$stable" == "1" ]]; then
  extensions_root="$HOME/.vscode/extensions"
else
  extensions_root="$HOME/.vscode-insiders/extensions"
fi

mkdir -p "$extensions_root"

for existing in "$extensions_root"/"$ext_id"*; do
  [[ -e "$existing" ]] || continue
  if [[ "$force" == "1" ]]; then
    rm -rf "$existing"
  else
    echo "Existing extension found: $existing"
    echo "Re-run with --force to replace it."
    exit 1
  fi
done

ln -s "$repo_root" "$extensions_root/$name"
echo "Installed dev link for $ext_id"
echo "Run 'npm run compile' after changes, then reload the VS Code window."
