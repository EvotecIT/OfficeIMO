#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../../.." && pwd)"
CAPTURE_ROOT="${CAPTURE_ROOT:-$ROOT_DIR/Artifacts/Studio/Apple/screenshots}"
OUTPUT_DIR="${OUTPUT_DIR:-$ROOT_DIR/Website/static/images/studio}"
MODE="check"
EXPECTED_SOURCE_COMMIT=""

usage() {
  cat <<'EOF'
Usage:
  Build/Studio/Apple/website-screenshot-promote.sh [--check | --apply] --source-commit <sha> [options]

Options:
  --check                 Validate and report changed website assets (default)
  --apply                 Copy the reviewed candidate into Website/static/images/studio
  --source-commit <sha>   Exact master commit represented by the capture
  --capture-root <path>   Restored exact capture root
  --output-dir <path>     Website screenshot asset directory
  --help                  Show this help

The capture provenance must contain the exact website candidate bytes. This
command never captures the app and never commits, pushes, or deploys the site.
EOF
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --check)
      MODE="check"
      shift
      ;;
    --apply)
      MODE="apply"
      shift
      ;;
    --capture-root)
      [[ $# -ge 2 ]] || { echo "--capture-root requires a path" >&2; exit 1; }
      CAPTURE_ROOT="$2"
      shift 2
      ;;
    --output-dir)
      [[ $# -ge 2 ]] || { echo "--output-dir requires a path" >&2; exit 1; }
      OUTPUT_DIR="$2"
      shift 2
      ;;
    --source-commit)
      [[ $# -ge 2 ]] || { echo "--source-commit requires a SHA" >&2; exit 1; }
      EXPECTED_SOURCE_COMMIT="$(printf '%s' "$2" | tr '[:upper:]' '[:lower:]')"
      shift 2
      ;;
    --help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown argument: $1" >&2
      usage
      exit 1
      ;;
  esac
done

[[ "$EXPECTED_SOURCE_COMMIT" =~ ^[0-9a-f]{40}$ ]] || {
  echo "--source-commit must be an exact 40-character Git commit SHA" >&2
  exit 1
}

CAPTURE_ROOT="$(cd "$CAPTURE_ROOT" && pwd -P)"
CANDIDATE_ROOT="$CAPTURE_ROOT/website-candidate"
CANDIDATE_SCREENSHOTS="$CANDIDATE_ROOT/screenshots"
PROVENANCE_PATH="$CAPTURE_ROOT/powerforge-apple-screenshot-provenance.json"
CONTACT_SHEET="$CANDIDATE_ROOT/contact-sheet.png"
EXPECTED=()
for family in iphone ipad mac; do
  for scene in dashboard markup form office; do
    EXPECTED+=("apple-$family-$scene.png")
  done
done

[[ -f "$PROVENANCE_PATH" && ! -L "$PROVENANCE_PATH" ]] || {
  echo "Missing unlinked capture provenance: $PROVENANCE_PATH" >&2
  exit 1
}
[[ -f "$CONTACT_SHEET" && ! -L "$CONTACT_SHEET" ]] || {
  echo "Missing unlinked website review contact sheet: $CONTACT_SHEET" >&2
  exit 1
}

ruby -rjson -rdigest -e '
  capture_root = File.realpath(ARGV.shift)
  provenance_path = ARGV.shift
  expected_source_commit = ARGV.shift
  expected = ARGV
  provenance = JSON.parse(File.read(provenance_path))
  abort("capture provenance source commit is invalid") unless provenance.fetch("sourceCommit").match?(/\A[0-9a-f]{40}\z/)
  abort("capture provenance does not match --source-commit") unless provenance.fetch("sourceCommit") == expected_source_commit
  abort("capture provenance is not from OfficeIMO master") unless provenance.fetch("repository") == "EvotecIT/OfficeIMO" &&
    provenance.fetch("workflowRef") == "EvotecIT/OfficeIMO/.github/workflows/studio-apple-capture.yml@refs/heads/master"
  inventory = provenance.fetch("screenshots").to_h { |item| [item.fetch("path"), item.fetch("sha256").downcase] }
  candidate_root = File.join(capture_root, "website-candidate", "screenshots")
  actual_candidates = Dir.glob(File.join(candidate_root, "**", "*.png"), File::FNM_DOTMATCH)
    .select { |path| File.file?(path) }
    .map { |path| path.delete_prefix(capture_root + File::SEPARATOR) }
    .sort
  expected_candidates = expected.drop(1).sort
  abort("website candidate must contain exactly the expected screenshots") unless actual_candidates == expected_candidates
  expected.each do |relative|
    joined = File.join(capture_root, relative)
    stat = File.lstat(joined)
    abort("candidate must not be a symbolic link: #{relative}") if stat.symlink?
    abort("candidate is not a regular file: #{relative}") unless stat.file?
    full = File.realpath(joined)
    abort("candidate escapes capture root: #{relative}") unless full.start_with?(capture_root + File::SEPARATOR)
    actual = Digest::SHA256.file(full).hexdigest
    abort("candidate is absent from capture provenance: #{relative}") unless inventory.fetch(relative, nil) == actual
  end
  puts "Verified OfficeIMO Studio #{provenance.fetch("marketingVersion")} website candidates from #{expected_source_commit}."
' "$CAPTURE_ROOT" "$PROVENANCE_PATH" \
  "$EXPECTED_SOURCE_COMMIT" \
  "website-candidate/contact-sheet.png" \
  "${EXPECTED[@]/#/website-candidate/screenshots/}"

changed=0
for relative in "${EXPECTED[@]}"; do
  source="$CANDIDATE_SCREENSHOTS/$relative"
  destination="$OUTPUT_DIR/$relative"
  [[ -f "$source" && ! -L "$source" ]] || {
    echo "Missing unlinked website candidate: $source" >&2
    exit 1
  }
  if [[ ! -f "$destination" ]] || ! cmp -s "$source" "$destination"; then
    echo "CHANGED $relative"
    changed=$((changed + 1))
  fi
done

if [[ "$MODE" == "check" ]]; then
  echo "$changed website screenshot(s) would change. Review: $CONTACT_SHEET"
  exit 0
fi

mkdir -p "$OUTPUT_DIR"
for relative in "${EXPECTED[@]}"; do
  destination="$OUTPUT_DIR/$relative"
  [[ ! -L "$destination" ]] || {
    echo "Refusing to replace linked website asset: $destination" >&2
    exit 1
  }
  cp "$CANDIDATE_SCREENSHOTS/$relative" "$destination"
done
echo "Promoted $changed reviewed website screenshot(s) into $OUTPUT_DIR"
