#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$repo_root"

if [[ ! -d node_modules ]]; then
  npm install --include=dev
fi

dotnet build ../OfficeIMO.Markup.Cli/OfficeIMO.Markup.Cli.csproj -c Release --framework net8.0 --no-restore -m:1 -p:BuildInParallel=false -p:UseSharedCompilation=false --nologo --verbosity minimal
rm -rf "$repo_root/tools/OfficeIMO.Markup.Cli"
mkdir -p "$repo_root/tools/OfficeIMO.Markup.Cli"
cp -R "$repo_root/../OfficeIMO.Markup.Cli/bin/Release/net8.0/." "$repo_root/tools/OfficeIMO.Markup.Cli/"
npm run compile
npx vsce package --allow-missing-repository

vsix_name="$(node -e "const p=require('./package.json'); console.log(p.name + '-' + p.version + '.vsix')")"
code-insiders --install-extension "$repo_root/$vsix_name" "${1:-}"
echo "Installed. Reload VS Code Insiders to activate the update."
