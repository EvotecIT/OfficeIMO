"""Compare paired CodeQL extraction inventories and full-suite SARIF results.

Inputs are CSV from ExtractedFiles.ql and SARIF from the same CLI/query bundle.
Run against unchanged product source and output paths; do not normalize away
source, generated-file, location, query or message differences.
"""
import argparse
import collections
import csv
import json
from pathlib import Path


def inventory(path):
    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.reader(stream)
        next(reader)  # CodeQL column labels.
        return collections.Counter(tuple(row) for row in reader)


def analysis(path):
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    if not data.get("runs"):
        raise ValueError(f"{path}: no completed CodeQL analysis run")
    rules = set()
    queries = set()
    results = collections.Counter()
    for run in data["runs"]:
        if run["tool"]["driver"].get("name") != "CodeQL" or not run["tool"]["driver"].get("rules"):
            raise ValueError(f"{path}: missing CodeQL query inventory")
        invocations = run.get("invocations", [])
        if not invocations or any(invocation.get("executionSuccessful") is not True for invocation in invocations):
            raise ValueError(f"{path}: analysis did not complete successfully")
        for invocation in invocations:
            notifications = invocation.get("toolExecutionNotifications", []) + invocation.get("toolConfigurationNotifications", [])
            if any(notification.get("level") == "error" for notification in notifications):
                raise ValueError(f"{path}: analysis contains an error notification")
        driver = run["tool"]["driver"]
        rules.update(rule["id"] for rule in driver["rules"])
        # Diagnostic queries are SARIF notification descriptors; metrics and
        # alert queries are rules. All belong to the resolved full suite.
        queries.update(item["id"] for group in ("rules", "notifications") for item in driver.get(group, []))
        for result in run.get("results", []):
            # Tool timestamps, database paths and invocation timing are not alerts.
            evidence = {key: result[key] for key in ("ruleId", "level", "message", "locations", "codeFlows") if key in result}
            results[json.dumps(evidence, sort_keys=True)] += 1
    return rules, queries, results


def compare(before_files, after_files, before_sarif, after_sarif, expected_rules):
    before = inventory(before_files)
    after = inventory(after_files)
    before_rules, before_queries, before_results = analysis(before_sarif)
    after_rules, after_queries, after_results = analysis(after_sarif)
    expected = set(json.loads(expected_rules.read_text(encoding="utf-8-sig")))
    if not expected:
        raise ValueError("The recorded full query-suite expectation is empty")
    report = {
        "equal": bool(before) and before == after and before_rules == after_rules and before_results == after_results
                 and expected <= before_queries and expected <= after_queries,
        "before_files": sum(before.values()), "after_files": sum(after.values()),
        "missing_files": list((before - after).elements()), "added_files": list((after - before).elements()),
        "before_rules": len(before_rules), "after_rules": len(after_rules),
        "missing_rules": sorted(before_rules - after_rules), "added_rules": sorted(after_rules - before_rules),
        "expected_queries": len(expected), "before_queries": len(expected & before_queries), "after_queries": len(expected & after_queries),
        "unproven_queries": sorted(expected - before_queries | expected - after_queries),
        "before_results": sum(before_results.values()), "after_results": sum(after_results.values()),
        "missing_results": list((before_results - after_results).elements()),
        "added_results": list((after_results - before_results).elements()),
    }
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("before_files", type=Path)
    parser.add_argument("after_files", type=Path)
    parser.add_argument("before_sarif", type=Path)
    parser.add_argument("after_sarif", type=Path)
    parser.add_argument("expected_rules", type=Path, help="JSON query IDs resolved from the unchanged full suite")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    report = compare(args.before_files, args.after_files, args.before_sarif, args.after_sarif, args.expected_rules)
    args.output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: value for key, value in report.items() if not isinstance(value, list)}, indent=2))
    raise SystemExit(0 if report["equal"] else 1)


if __name__ == "__main__":
    main()
