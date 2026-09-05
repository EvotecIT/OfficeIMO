#!/usr/bin/env bash
set -euo pipefail

# An opt-in Linux/WSL acceptance environment; no system mount or real disk is filled.
script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
assembly=$(readlink -f -- "${1:-$script_dir/bin/Debug/net10.0/OfficeIMO.Studio.Tests.dll}")
test -f "$assembly"
task_root=$(mktemp -d /tmp/officeimo-storage-XXXXXXXX)
cleanup() {
    case "$task_root" in
        /tmp/officeimo-storage-*) rm -rf -- "$task_root" ;;
        *) return 1 ;;
    esac
}
trap cleanup EXIT
mkdir -- "$task_root/volume" "$task_root/backing"
export OFFICEIMO_STUDIO_PARENT_MOUNT_NAMESPACE=$(readlink /proc/self/ns/mnt)
unshare --user --map-root-user --mount bash -euc '
    mount --make-rprivate /
    mount -t tmpfs -o size=8m,mode=700 tmpfs "$1/volume"
    dotnet "$2" --studio-process-probe storage-full "$1"
    umount "$1/volume"
    mount -t tmpfs -o size=8m,mode=700 tmpfs "$1/backing"
    mount --bind "$1/backing" "$1/volume"
    dotnet "$2" --studio-process-probe storage-detach "$1"
    umount "$1/backing"
' storage-acceptance "$task_root" "$assembly"
