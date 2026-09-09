import sys
import csv
import json
import statistics
from collections import defaultdict
from pathlib import Path

root = Path(__file__).resolve().parent
label = sys.argv[1]
aa = label == "aa-calibration"
repeats = 2 if aa else 3
exports = 16 if aa else 8
run_names = [f"{label}-r{repeat}-{mask}" for repeat in range(1, repeats + 1) for mask in (65535, 4294901760)]
output = {"protocol": {"iterations": 30, "warmups": 12, "exportsPerSample": exports, "order": "Rotated", "outliersRemoved": False}, "runs": []}
all_values = defaultdict(lambda: {"Before": [], "After": [], "runChanges": []})
processes = set()
for run_name in run_names:
    files = list((root / run_name).glob("*/samples.csv"))
    assert len(files) == 1, (run_name, "missing or ambiguous samples")
    folder = files[0].parent
    metadata = json.loads((folder / "metadata.json").read_text(encoding="utf-8-sig"))
    assert metadata["benchmark.ProcessId"] not in processes
    processes.add(metadata["benchmark.ProcessId"])
    assert metadata["benchmark.ExportsPerOperation"] == str(exports)
    assert metadata["benchmark.ComparisonMode"] == ("Identical baseline assemblies" if aa else "Baseline versus candidate")
    with files[0].open(encoding="utf-8-sig", newline="") as stream:
        rows = list(csv.DictReader(stream))
    grouped = defaultdict(lambda: {"Before": {}, "After": {}})
    for row in rows:
        assert row["Status"] == "Succeeded"
        grouped[row["Scenario"]][row["Engine"]][int(row["Iteration"])] = float(row["DurationMs"]) / exports
    run = {"id": folder.name, "label": run_name, "mask": metadata["benchmark.AffinityMask"], "processId": metadata["benchmark.ProcessId"], "cases": {}}
    for scenario, engines in sorted(grouped.items()):
        assert all(set(samples) == set(range(30)) for samples in engines.values())
        before, after = ([engines[engine][i] for i in range(30)] for engine in ("Before", "After"))
        means = statistics.mean(before), statistics.mean(after)
        medians = statistics.median(before), statistics.median(after)
        mean_change = (means[1] / means[0] - 1) * 100
        median_change = (medians[1] / medians[0] - 1) * 100
        run["cases"][scenario] = {"meanMs": means, "medianMs": medians, "meanChangePercent": mean_change, "medianChangePercent": median_change, "pairedSamplesMs": list(zip(before, after))}
        aggregate = all_values[(scenario, run["mask"])]
        aggregate["Before"].extend(before)
        aggregate["After"].extend(after)
        aggregate["runChanges"].append((mean_change, median_change))
    assert len(run["cases"]) == (7 if aa else 22)
    output["runs"].append(run)

aggregates = []
for (scenario, mask), values in sorted(all_values.items()):
    means = [statistics.mean(values[engine]) for engine in ("Before", "After")]
    medians = [statistics.median(values[engine]) for engine in ("Before", "After")]
    changes = values["runChanges"]
    assert len(changes) == repeats
    aggregate = {"scenario": scenario, "mask": mask, "meanMs": means, "medianMs": medians,
                 "meanChangePercent": (means[1] / means[0] - 1) * 100,
                 "medianChangePercent": (medians[1] / medians[0] - 1) * 100,
                 "perRunMeanChanges": [change[0] for change in changes],
                 "perRunMedianChanges": [change[1] for change in changes]}
    aggregates.append(aggregate)
    print(f"{scenario:30} {mask:10} mean {aggregate['meanChangePercent']:+6.1f}% median {aggregate['medianChangePercent']:+6.1f}% run means " + ", ".join(f"{value:+.1f}" for value in aggregate["perRunMeanChanges"]))
output["aggregated"] = aggregates
(root / (label + "-results.json")).write_text(json.dumps(output, separators=(",", ":")), encoding="utf-8")
