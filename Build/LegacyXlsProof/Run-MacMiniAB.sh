#!/bin/bash
# Run on evomini-agents with the same benchmark sources in both revision trees.
set -euo pipefail
if [ "$#" -ne 4 ]; then
    printf 'Usage: %s MASTER_CHECKOUT CANDIDATE_CHECKOUT MASTER_CAPTURE OUTPUT\n' "$0" >&2
    exit 2
fi

master=$(cd "$1" && pwd -P)
candidate=$(cd "$2" && pwd -P)
capture=$(cd "$3" && pwd -P)
mkdir -p "$4"
output=$(cd "$4" && pwd -P)
lock=/Volumes/Samsung/Development/Scratch/officeimo-bench.lock
project=Build/LegacyXlsProof/Benchmarks/Benchmarks.csproj
probe=$(dirname "$project")
diff -q "$master/$project" "$candidate/$project"
for source in Program.cs PlainImportBenchmarks.cs; do
    diff -q "$master/$probe/$source" "$candidate/$probe/$source"
done

exec lockf -k -t 900 "$lock" /bin/bash -s -- "$master" "$candidate" "$capture" "$output" "$project" <<'RUN_LOCKED'
set -euo pipefail
master=$1
candidate=$2
capture=$3
output=$4
project=$5
for result in master-first-call.csv candidate-first-call.csv master candidate benchmark-order.txt; do
    if [[ -e "$output/$result" || -L "$output/$result" ]]; then
        printf 'Use a fresh evidence directory. Existing result: %s\n' "$output/$result" >&2
        exit 2
    fi
done
mkdir -p "$output/tmp"
export TMPDIR="$output/tmp" TZ=UTC OFFICEIMO_LEGACY_XLS_PROOF_OUTPUT="$capture"

for checkout in "$master" "$candidate"; do
    (cd "$checkout" && dotnet build "$project" -c Release -m:1)
done

wait_for_quiet() {
    local deadline=$((SECONDS + 900))
    while [ "$SECONDS" -lt "$deadline" ]; do
        if sysctl -n vm.loadavg | awk '{exit !($2 < 3.0)}'; then
            return
        fi
        sleep 15
    done
    printf 'The Mac mini did not become quiet within 15 minutes.\n' >&2
    return 1
}

sysctl -n machdep.cpu.brand_string
sw_vers
dotnet --version
measure_first_call() {
    local checkout=$1
    local label=$2
    dotnet "$checkout/$(dirname "$project")/bin/Release/net10.0/Benchmarks.dll" --startup "${OFFICEIMO_LEGACY_XLS_STARTUP_WORKLOAD:-stream}" >> "$output/$label-first-call.csv"
}

wait_for_quiet
for attempt in 1 2 3 4 5; do
    if (( attempt % 2 == 1 )); then
        measure_first_call "$master" master
        measure_first_call "$candidate" candidate
    else
        measure_first_call "$candidate" candidate
        measure_first_call "$master" master
    fi
done

run() {
    local checkout=$1
    local label=$2
    wait_for_quiet
    uptime
    wc -c "$checkout/OfficeIMO.Excel/bin/Release/net10.0/OfficeIMO.Excel.dll" > "$output/$label-excel-assembly.txt"
    (cd "$checkout" && dotnet run --project "$project" -c Release --no-build -- \
        --filter "${OFFICEIMO_LEGACY_XLS_BENCHMARK_FILTER:-LegacyXlsStreamingBenchmarks.*}" --exporters json --artifacts "$output/$label")
}
if (( RANDOM % 2 )); then
    printf 'master\ncandidate\n' > "$output/benchmark-order.txt"
    run "$master" master
    run "$candidate" candidate
else
    printf 'candidate\nmaster\n' > "$output/benchmark-order.txt"
    run "$candidate" candidate
    run "$master" master
fi
uptime
RUN_LOCKED
