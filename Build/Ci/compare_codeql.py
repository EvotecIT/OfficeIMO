"""Compare CodeQL inventories, SARIF alerts and full-suite execution evidence.

Inputs are CSV from ExtractedFiles.ql and SARIF from the same CLI/query bundle.
Run against unchanged product source and output paths; do not normalize away
source, generated-file, alert-location, query-identity or alert-message differences.

Equality covers the extracted CSV inventory, rule and executed-query ID sets,
and projected SARIF alerts. BQRS artifacts attest execution and ZIP integrity;
their rows and diagnostic notification messages are not compared.
"""
import argparse
import collections
import csv
import json
import re
import zipfile
from pathlib import Path


def inventory(path):
    """Read the extracted path/line-count multiset, retaining duplicates."""
    with path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.reader(stream)
        next(reader)  # CodeQL column labels.
        return collections.Counter(tuple(row) for row in reader)


def analysis(path):
    """Validate CodeQL SARIF and retain query descriptors and alert evidence."""
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    if not data.get("runs"):
        raise ValueError(f"{path}: no completed CodeQL analysis run")
    rules = set()
    queries = set()
    results = collections.Counter()
    for run in data["runs"]:
        if run["tool"]["driver"].get("name") not in {"CodeQL", "CodeQL command-line toolchain"} or not run["tool"]["driver"].get("rules"):
            raise ValueError(f"{path}: missing CodeQL query inventory")
        invocations = run.get("invocations", [])
        if any(invocation.get("executionSuccessful") is not True for invocation in invocations):
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
            evidence = {key: result[key] for key in ("ruleId", "level", "message", "locations", "relatedLocations", "codeFlows") if key in result}
            results[json.dumps(evidence, sort_keys=True)] += 1
    return rules, queries, results


def executed_queries(bundle, metadata, expected):
    """Require each resolved query's BQRS in a clean, fully analyzed database ZIP.

    SARIF descriptors can include unexecuted rules. BQRS files are the query
    evaluator's outputs, retained by database bundle --include-results.
    """
    resolved = json.loads(metadata.read_text(encoding="utf-8-sig"))
    if len(resolved) != len(expected) or {query["id"] for query in resolved} != expected:
        raise ValueError("Resolved query metadata must cover the exact expected suite")
    executed = set()
    with zipfile.ZipFile(bundle) as archive:
        for query in resolved:
            match = re.search(r"/codeql/csharp-queries/[^/]+/(.+)\.ql$", query["query"].replace("\\", "/"))
            if not match:
                raise ValueError(f"Cannot locate C# query result for {query['id']}")
            suffix = "/results/codeql/csharp-queries/" + match[1] + ".bqrs"
            entries = [entry for entry in archive.infolist() if ("/" + entry.filename).endswith(suffix)]
            if len(entries) == 1 and entries[0].file_size > 0:
                with archive.open(entries[0]) as result:
                    # Reading the entire result verifies its ZIP CRC without
                    # depending on an undocumented BQRS binary layout.
                    while result.read(65536):
                        pass
                executed.add(query["id"])
    return executed


def compare(before_files, after_files, before_sarif, after_sarif, expected_rules,
            query_metadata, before_query_results, after_query_results):
    """Compare inventory, query execution IDs and SARIF alerts, excluding BQRS rows."""
    before = inventory(before_files)
    after = inventory(after_files)
    before_rules, before_queries, before_results = analysis(before_sarif)
    after_rules, after_queries, after_results = analysis(after_sarif)
    expected = set(json.loads(expected_rules.read_text(encoding="utf-8-sig")))
    if not expected:
        raise ValueError("The recorded full query-suite expectation is empty")
    before_executed = executed_queries(before_query_results, query_metadata, expected)
    after_executed = executed_queries(after_query_results, query_metadata, expected)
    report = {
        "equal": bool(before) and before == after and before_rules == after_rules and before_results == after_results
                 and expected <= before_queries and expected <= after_queries
                 and expected == before_executed == after_executed,
        "before_files": sum(before.values()), "after_files": sum(after.values()),
        "missing_files": list((before - after).elements()), "added_files": list((after - before).elements()),
        "before_rules": len(before_rules), "after_rules": len(after_rules),
        "missing_rules": sorted(before_rules - after_rules), "added_rules": sorted(after_rules - before_rules),
        "expected_queries": len(expected), "before_queries": len(expected & before_queries), "after_queries": len(expected & after_queries),
        "before_executed_queries": len(before_executed), "after_executed_queries": len(after_executed),
        "unproven_queries": sorted(expected - before_queries | expected - after_queries
                                   | expected - before_executed | expected - after_executed),
        "before_results": sum(before_results.values()), "after_results": sum(after_results.values()),
        "missing_results": list((before_results - after_results).elements()),
        "added_results": list((after_results - before_results).elements()),
    }
    return report


def main():
    """Write a machine-readable comparison and fail on missing or changed proof."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("before_files", type=Path)
    parser.add_argument("after_files", type=Path)
    parser.add_argument("before_sarif", type=Path)
    parser.add_argument("after_sarif", type=Path)
    parser.add_argument("expected_rules", type=Path, help="JSON query IDs resolved from the unchanged full suite")
    parser.add_argument("--query-metadata", type=Path, required=True, help="Resolved IDs and absolute query paths from the unchanged C# query pack")
    parser.add_argument("--before-query-results", type=Path, required=True, help="Baseline database bundle ZIP with results included")
    parser.add_argument("--after-query-results", type=Path, required=True, help="Candidate database bundle ZIP with results included")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    report = compare(args.before_files, args.after_files, args.before_sarif, args.after_sarif, args.expected_rules,
                     args.query_metadata, args.before_query_results, args.after_query_results)
    args.output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({key: value for key, value in report.items() if not isinstance(value, list)}, indent=2))
    raise SystemExit(0 if report["equal"] else 1)


if __name__ == "__main__":
    main()
