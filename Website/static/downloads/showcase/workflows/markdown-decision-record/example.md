# ADR-007: Keep report inputs with the output

Status: Accepted \| Owner: Reporting team \| Date: 14 September 2026

## Context

A generated PDF is easy to share, but reviewers also need to understand which data and configuration produced it.

## Decision

Deliver the report, its structured input snapshot, and a small manifest as one versioned bundle.

```text
report-bundle/
  report.pdf
  input.json
  manifest.json
```

## Alternatives considered

| Alternative | Reason not selected |
| --- | --- |
| PDF only | Insufficient context for reproducing the report |
| Live dashboard link only | The underlying data can change after review |
| Full database export | Unnecessary volume and unrelated information |

## Consequences

- The input snapshot must contain only the fields used by the report.
- The manifest records the generator version and file hashes.
- Retention rules apply to the whole bundle.

## Revisit when

The report contains data that cannot be retained, or the generation volume makes the bundle impractical.
