#!/usr/bin/env bash
set -euo pipefail

if [[ "$#" -eq 0 ]]; then
  echo "At least one Ubuntu package is required."
  exit 2
fi

# GitHub-hosted Ubuntu mirrors occasionally accept a request and then stop
# transferring package indexes. Bound each attempt and let apt retry instead of
# consuming the entire interoperability-job budget on one stalled mirror read.
apt_options=(
  -o Acquire::Retries=3
  -o Acquire::http::Timeout=30
  -o Acquire::https::Timeout=30
)

sudo apt-get "${apt_options[@]}" update
sudo env DEBIAN_FRONTEND=noninteractive \
  apt-get "${apt_options[@]}" install --yes --no-install-recommends "$@"
