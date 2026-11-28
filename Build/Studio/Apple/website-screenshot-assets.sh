#!/usr/bin/env bash
# Derives reviewable website screenshots from the English raw store captures.
# They stay in the capture until website-screenshot-promote.sh copies them.
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../../.." && pwd)"
CAPTURE_ROOT="${1:-$ROOT_DIR/Artifacts/Studio/Apple/screenshots}"
OUTPUT_DIR="${2:-$CAPTURE_ROOT/website-candidate/screenshots}"
# The scenes every family captures, as website-screenshot-promote.sh expects them.
SCENES=(dashboard markup form office)

render_image() {
  local input_path="$1"
  local width="$2"
  local output_path="$3"
  [[ -s "$input_path" ]] || {
    echo "Missing screenshot: $input_path" >&2
    exit 1
  }
  sips --resampleWidth "$width" "$input_path" --out "$output_path" >/dev/null
}

mkdir -p "$OUTPUT_DIR"
for scene in "${SCENES[@]}"; do
  render_image "$CAPTURE_ROOT/iPhone/en-US/studio-$scene-iphone.png" 720 "$OUTPUT_DIR/apple-iphone-$scene.png"
  render_image "$CAPTURE_ROOT/iPad/en-US/studio-$scene-ipad.png" 1200 "$OUTPUT_DIR/apple-ipad-$scene.png"
  render_image "$CAPTURE_ROOT/macOS/en-US/studio-$scene-macos.png" 1440 "$OUTPUT_DIR/apple-mac-$scene.png"
done

echo "Website screenshots written under $OUTPUT_DIR"
