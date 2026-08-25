# OfficeIMO.Adf performance evidence (2026-08-24)

## Result

OfficeIMO.Adf is within the 2x contender margin for typed Atlas Document Format
JSON parsing and semantic parse/write round trips. Against a benchmark-only
`System.Text.Json` model that materializes the same document, node, mark,
attribute, and extension-data shape, OfficeIMO measures 0.98-1.18x elapsed time
and 1.24-1.27x managed allocation for parsing. Round trips measure 0.99-1.02x
elapsed time and 1.52-1.65x allocation.

The optimized parser retains unknown root, node, mark, and attribute values,
keeps public collections mutable, and validates the ADF structural contract.
All generated output is semantically identical to the input and has the same
UTF-8 byte count. The raw JSON-tree lane remains visible as a narrower platform
lower bound, but it is not used as the contender comparison because it does not
materialize an equivalent typed model.

## Comparison contract

The deterministic 25- and 500-paragraph corpora include headings, marked and
linked text, nested lists, tables, attributes, unknown node and mark fields, an
unknown node type, and root extension data. Every parse lane materializes the
complete corpus. Every round-trip lane parses and writes the complete JSON tree.

Preflight requires deep semantic JSON equality, preservation of the root
extension and unknown-node payload, and matching recursive node and text
counts. OfficeIMO additionally runs its structural ADF validator. The typed
platform model preserves the same modeled and extension-data shape but does not
perform ADF validation, so it remains a cost floor rather than a full product
competitor.

No accepted managed package was available for a license-independent,
feature-equivalent raw ADF parse, preserve, validate, and write comparison. A
commercial conversion package was not added to runtime or benchmark restore
without a configured license and an equivalent contract.

## BenchmarkDotNet results

Environment: Windows build 26200, AMD Ryzen 9 9950X3D2, .NET SDK 10.0.111,
.NET 8.0.30 x64, BenchmarkDotNet 0.15.8 `ShortRun`. The clean measured commit
was `5ec8cfaaa7e71538fb06a263fae6a4e45617329d`.

Ratios are OfficeIMO divided by the typed `System.Text.Json` model. Lower is
better.

| Workload | Scale | OfficeIMO mean | Typed floor mean | Time ratio | OfficeIMO allocation | Typed floor allocation | Allocation ratio |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Parse | Small, 11,516 B | 111.50 us | 94.56 us | 1.18x | 112.42 KiB | 90.41 KiB | 1.24x |
| Parse | Normal, 224,480 B | 2.083 ms | 2.135 ms | 0.98x | 2,164.49 KiB | 1,704.66 KiB | 1.27x |
| Round trip | Small, 11,516 B | 147.80 us | 145.48 us | 1.02x | 175.13 KiB | 114.96 KiB | 1.52x |
| Round trip | Normal, 224,480 B | 3.466 ms | 3.518 ms | 0.99x | 3,542.38 KiB | 2,143.31 KiB | 1.65x |

The raw JSON-tree floor is materially narrower. OfficeIMO remains about 2.3x
that lane for parse allocation and about 3.2-3.3x for round-trip allocation;
those ratios describe the cost of materializing a typed preserving graph, not
an equivalent-product deficit.

## Isolated peak-memory and size evidence

Each workload, scale, implementation, and repetition ran in a fresh child
process. Values are ratios of three-run medians from clean commit
`c264ac20b24808422fef79cfe21f601a7d7d5749`.

| Workload | Scale | Time ratio | Allocation ratio | Retained-heap ratio | Managed-peak ratio | Process-peak ratio | Input/output bytes |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Parse | Small | 1.56x | 1.24x | 1.01x | 1.24x | 0.93x | 11,516 / not applicable |
| Parse | Normal | 2.02x | 1.27x | 0.99x | 1.27x | 1.02x | 224,480 / not applicable |
| Round trip | Small | 1.88x | 1.52x | 1.11x | 1.51x | 0.94x | 11,516 / 11,516 |
| Round trip | Normal | 1.89x | 1.65x | 1.00x | 1.12x | 0.98x | 224,480 / 224,480 |

The isolated elapsed values are directional because process scheduling and
tiered compilation remain visible; the controlled BenchmarkDotNet result is the
throughput classification. The normal parse median is 2.02x in isolated
processes but 0.98x in the controlled run. All allocation and peak-memory
ratios remain below 2x in both measurement styles.

## Optimization

The initial implementation eagerly allocated content, mark, attribute, and
extension collections for every node. It also rebuilt known-property sets and
full JSON error-path strings throughout successful parses, repeatedly scanned
each node object, and copied the serialized memory stream before UTF-8 string
creation.

The optimized implementation uses lazy mutable collections, allocation-free
internal empty views, one-pass node and mark property parsing, reusable property
classification, a mutable error-path builder, and direct extraction from the
serialization buffer. The public mutable model and exact unknown-data
preservation contracts remain covered on net472, net8.0, and net10.0.

## Reproduce

```powershell
dotnet test .\OfficeIMO.Adf.Tests\OfficeIMO.Adf.Tests.csproj -c Release
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Adf.Benchmarks -- --filter '*' --job short --artifacts .benchmark-artifacts\adf\bdn
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Adf.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\adf\isolated.json
```

The benchmark project stays outside the normal solution and uses only platform
JSON support plus the OfficeIMO project reference. Raw reports remain ignored
machine-local artifacts.
