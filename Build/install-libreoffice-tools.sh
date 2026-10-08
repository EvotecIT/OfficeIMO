#!/usr/bin/env bash
set -euo pipefail

if [[ "$#" -eq 0 ]]; then
  echo "At least one Ubuntu package is required."
  exit 2
fi

# Socket timeouts cover stalled reads; acquisition deadlines also cover mirrors
# that keep transferring packages too slowly for the interoperability-job budget.
apt_options=(
  -o Acquire::Retries=3
  -o Acquire::http::Timeout=30
  -o Acquire::https::Timeout=30
)

mirror_list=/etc/apt/apt-mirrors.txt
mirror_backup=''
workspace=''

restore_mirror_list() {
  if [[ -n "$mirror_backup" ]]; then
    sudo cp -- "$mirror_backup" "$mirror_list" || return "$?"
    mirror_backup=''
  fi
}

cleanup() {
  local status=$?
  trap - EXIT
  if ! restore_mirror_list; then
    echo "Could not restore the apt mirror list; original retained at $mirror_backup." >&2
    exit 1
  fi
  if [[ -n "$workspace" ]]; then
    rm -rf -- "$workspace"
  fi
  exit "$status"
}

trap cleanup EXIT
trap 'exit 130' INT
trap 'exit 143' TERM

acquire_packages() {
  sudo timeout --signal=TERM --kill-after=30s 2m \
    apt-get "${apt_options[@]}" update || return "$?"
  sudo timeout --signal=TERM --kill-after=30s 8m \
    env DEBIAN_FRONTEND=noninteractive apt-get "${apt_options[@]}" \
    install --download-only --yes --no-install-recommends "$@"
}

if acquire_packages "$@"; then
  echo "LibreOffice packages acquired from the configured Ubuntu sources."
else
  acquisition_status=$?
  echo "Ubuntu package acquisition failed or exceeded its deadline (exit $acquisition_status)." >&2
  if [[ "${GITHUB_ACTIONS:-}" != true || ! -f "$mirror_list" || -L "$mirror_list" || ! -r "$mirror_list" ]]; then
    echo "No recognized GitHub-hosted Ubuntu mirror list is available for a retry." >&2
    exit "$acquisition_status"
  fi

  workspace=$(mktemp -d "${RUNNER_TEMP:-${TMPDIR:-/tmp}}/officeimo-apt.XXXXXX")
  # Reuse the runner's existing alternatives and priority metadata. No apt source
  # declarations, signing keys, or unrelated repositories are changed.
  if ! awk '
    $1 == "http://azure.archive.ubuntu.com/ubuntu/" || $1 == "https://azure.archive.ubuntu.com/ubuntu/" { azure = 1; next }
    $1 == "http://archive.ubuntu.com/ubuntu/" || $1 == "https://archive.ubuntu.com/ubuntu/" ||
      $1 == "http://security.ubuntu.com/ubuntu/" || $1 == "https://security.ubuntu.com/ubuntu/" { alternate = 1 }
    { print }
    END { if (!azure || !alternate) exit 1 }
  ' "$mirror_list" > "$workspace/fallback"; then
    echo "The runner mirror list has no Azure entry with an existing Ubuntu alternative; no retry made." >&2
    exit "$acquisition_status"
  fi

  cp -- "$mirror_list" "$workspace/original"
  mirror_backup="$workspace/original"
  sudo cp -- "$workspace/fallback" "$mirror_list"
  echo "Retrying bounded package acquisition through the runner's existing Ubuntu alternatives."
  if acquire_packages "$@"; then
    restore_mirror_list
  else
    acquisition_status=$?
    echo "Ubuntu alternate-mirror acquisition failed (exit $acquisition_status); installation was not started." >&2
    exit "$acquisition_status"
  fi
fi

# Download deadlines never interrupt dpkg. Installation uses only the archives
# acquired above, and the original mirror configuration is already restored.
sudo env DEBIAN_FRONTEND=noninteractive \
  apt-get "${apt_options[@]}" install --no-download --yes --no-install-recommends "$@"
