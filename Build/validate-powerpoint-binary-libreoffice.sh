#!/usr/bin/env bash
set -euo pipefail

artifact_dir="${1:?Binary PowerPoint artifact directory is required}"
output_dir="${2:?LibreOffice output directory is required}"
soffice_bin="${3:-soffice}"

command -v "$soffice_bin" >/dev/null 2>&1 || {
  echo "LibreOffice executable '$soffice_bin' was not found."
  exit 1
}

pdf_dir="$output_dir/pdf"
roundtrip_dir="$output_dir/roundtrip"
mkdir -p "$pdf_dir" "$roundtrip_dir"
profile_dir="$(mktemp -d "${TMPDIR:-/tmp}/officeimo-ppt-libreoffice.XXXXXX")"
trap 'rm -rf "$profile_dir"' EXIT

source_count=0
while IFS= read -r -d '' source_path; do
  source_count=$((source_count + 1))
  base_name="$(basename "${source_path%.*}")"

  "$soffice_bin" --headless "-env:UserInstallation=file://$profile_dir" \
    --convert-to pdf --outdir "$pdf_dir" "$source_path"
  pdf_path="$pdf_dir/$base_name.pdf"
  test -s "$pdf_path" || {
    echo "LibreOffice did not render '$source_path' to a non-empty PDF."
    exit 1
  }

  "$soffice_bin" --headless "-env:UserInstallation=file://$profile_dir" \
    --convert-to ppt --outdir "$roundtrip_dir" "$source_path"
  ppt_path="$roundtrip_dir/$base_name.ppt"
  test -s "$ppt_path" || {
    echo "LibreOffice did not resave '$source_path' as a non-empty PPT file."
    exit 1
  }
done < <(find "$artifact_dir" -maxdepth 1 -type f \
  \( -name '*.ppt' -o -name '*.pot' -o -name '*.pps' \) -print0)

test "$source_count" -eq 3 || {
  echo "Expected three PPT/POT/PPS artifacts but found $source_count."
  exit 1
}
echo "LibreOffice opened, rendered, and resaved all $source_count binary PowerPoint variants."
