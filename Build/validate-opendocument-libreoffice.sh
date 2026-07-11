#!/usr/bin/env bash
set -euo pipefail

artifact_dir="${1:?OpenDocument artifact directory is required}"
output_dir="${2:?LibreOffice output directory is required}"
soffice_bin="${3:-soffice}"

command -v "$soffice_bin" >/dev/null 2>&1 || {
  echo "LibreOffice executable '$soffice_bin' was not found."
  exit 1
}
command -v unzip >/dev/null 2>&1 || {
  echo "unzip is required for package verification."
  exit 1
}

mkdir -p "$output_dir"
profile_dir="$(mktemp -d "${TMPDIR:-/tmp}/officeimo-libreoffice-profile.XXXXXX")"
trap 'rm -rf "$profile_dir"' EXIT

source_count=0
output_count=0
while IFS= read -r -d '' source_path; do
  source_count=$((source_count + 1))
  extension="${source_path##*.}"
  extension_output="$output_dir/$extension"
  mkdir -p "$extension_output"
  "$soffice_bin" --headless "-env:UserInstallation=file://$profile_dir" \
    --convert-to "$extension" --outdir "$extension_output" "$source_path"
  output_path="$extension_output/$(basename "$source_path")"
  test -f "$output_path" || {
    echo "LibreOffice did not produce '$output_path'."
    exit 1
  }
  unzip -tq "$output_path" >/dev/null
  output_count=$((output_count + 1))
done < <(find "$artifact_dir" -type f \( -name '*.odt' -o -name '*.ods' -o -name '*.odp' \) -print0)

test "$source_count" -gt 0 || {
  echo "No ODT, ODS, or ODP packages were found in '$artifact_dir'."
  exit 1
}
test "$source_count" -eq "$output_count"
echo "LibreOffice opened and resaved $output_count OpenDocument packages."
