#!/usr/bin/env bash
set -euo pipefail

artifact_dir="${1:?artifact directory is required}"
schema_dir="${2:-${TMPDIR:-/tmp}/officeimo-odf-schema}"
schema14_url="https://docs.oasis-open.org/office/OpenDocument/v1.4/os/schemas/OpenDocument-v1.4-schema.rng"
manifest14_url="https://docs.oasis-open.org/office/OpenDocument/v1.4/os/schemas/OpenDocument-v1.4-manifest-schema.rng"
schema14_sha="4034ec6be29205d5fc1ee5f42468ac6ef824287b3aba6d9289032af4fafbda7f"
manifest14_sha="f96e1fe95f4e3609717a307c28e2fe0654c890fe70ea03eb8a53caa20b684782"
schema13_url="https://docs.oasis-open.org/office/OpenDocument/v1.3/os/schemas/OpenDocument-v1.3-schema.rng"
manifest13_url="https://docs.oasis-open.org/office/OpenDocument/v1.3/os/schemas/OpenDocument-v1.3-manifest-schema.rng"
schema13_sha="40bad03efdbb02825230d357da0aa6ac679934c5bf56c6281752c0c24d58e4e6"
manifest13_sha="8aee71f03484be112af972d622cc9031280c007b0a454ae8815e8e333c9bdd17"

mkdir -p "$schema_dir"
curl --fail --silent --show-error --location "$schema14_url" --output "$schema_dir/OpenDocument-v1.4-schema.rng"
curl --fail --silent --show-error --location "$manifest14_url" --output "$schema_dir/OpenDocument-v1.4-manifest-schema.rng"
curl --fail --silent --show-error --location "$schema13_url" --output "$schema_dir/OpenDocument-v1.3-schema.rng"
curl --fail --silent --show-error --location "$manifest13_url" --output "$schema_dir/OpenDocument-v1.3-manifest-schema.rng"
checksum() {
  if command -v sha256sum >/dev/null 2>&1; then
    sha256sum "$1" | awk '{print $1}'
  else
    shasum --algorithm 256 "$1" | awk '{print $1}'
  fi
}

test "$(checksum "$schema_dir/OpenDocument-v1.4-schema.rng")" = "$schema14_sha"
test "$(checksum "$schema_dir/OpenDocument-v1.4-manifest-schema.rng")" = "$manifest14_sha"
test "$(checksum "$schema_dir/OpenDocument-v1.3-schema.rng")" = "$schema13_sha"
test "$(checksum "$schema_dir/OpenDocument-v1.3-manifest-schema.rng")" = "$manifest13_sha"

for package in "$artifact_dir"/*.odt "$artifact_dir"/*.ods "$artifact_dir"/*.odp; do
  extract_dir="$schema_dir/$(basename "$package")"
  rm -rf "$extract_dir"
  mkdir -p "$extract_dir"
  unzip -q "$package" -d "$extract_dir"
  if [[ "$(basename "$package")" == *-1.3.* ]]; then
    document_schema="$schema_dir/OpenDocument-v1.3-schema.rng"
    manifest_schema="$schema_dir/OpenDocument-v1.3-manifest-schema.rng"
  else
    document_schema="$schema_dir/OpenDocument-v1.4-schema.rng"
    manifest_schema="$schema_dir/OpenDocument-v1.4-manifest-schema.rng"
  fi
  for part in content.xml styles.xml meta.xml settings.xml; do
    xmllint --noout --relaxng "$document_schema" "$extract_dir/$part"
  done
  xmllint --noout --relaxng "$manifest_schema" "$extract_dir/META-INF/manifest.xml"
done

for flat in "$artifact_dir"/*.fodt "$artifact_dir"/*.fods "$artifact_dir"/*.fodp; do
  xmllint --noout --relaxng "$schema_dir/OpenDocument-v1.4-schema.rng" "$flat"
done
