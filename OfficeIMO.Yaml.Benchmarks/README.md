# OfficeIMO YAML Benchmarks

This opt-in project measures OfficeIMO.Yaml and the reference YAML libraries, YamlDotNet 18.0.0 and SharpYaml 3.17.0, against the same conformance data and workloads. The libraries are referenced here only; OfficeIMO runtime projects do not depend on this project.

## Conformance comparison

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Yaml.Benchmarks -- --conformance --output yaml-conformance.md
```

The command runs each library through the harness in `OfficeIMO.Yaml.Tests/Conformance`, which this project compiles by link, and prints a Markdown report. Add `--details` to list every failing case and file, and `--timeout <seconds>` to change the per-call limit (default 10).

Inputs are the pinned [yaml-test-suite](../OfficeIMO.TestAssets/Yaml/yaml-test-suite/manifest.json) cases and the [real-world corpus](../OfficeIMO.TestAssets/Yaml/Corpus/manifest.json). Loading either fails if files drift from their manifests.

| Check | Passes when |
| --- | --- |
| Events | a valid case produces exactly the `test.event` lines, including scalar style, flow markers and explicit document markers |
| Error cases | an invalid case is rejected |
| JSON | a valid case with `in.json` loads to equal values; objects are unordered and numbers compare by value |
| Round-trip bytes | loading and saving without edits returns the input bytes |
| Round-trip content | re-parsing the saved output gives the same values, tags, anchors and aliases; style is ignored |
| Accepts | a corpus file parses |

Failures (F) are wrong output or accepted invalid input. Errors (E) are exceptions or timeouts on input the check expects to work.

Each adapter decodes bytes with YAML byte order mark and null-pattern detection and re-encodes saved text the same way, because both libraries read text rather than bytes. JSON loading uses each library's tree loader with one shared YAML 1.2 core-schema scalar resolver, so the JSON column compares structure loading rather than resolver choices. The SharpYaml syntax tree row reports only round-trips: the library documents it as reproducing its source text, and it has no event or value view.

## Throughput workloads

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Yaml.Benchmarks -- --validate-workloads
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Yaml.Benchmarks -- --filter "*YamlParseEventsBenchmarks*"
```

Four benchmark classes cover the operations, each with YamlDotNet as the baseline:

- `YamlParseEventsBenchmarks`: parse to events without building a tree. The OfficeIMO.Yaml lane reads every scalar value, alias name, anchor and tag as a string, matching the reference parsers' events, which carry them.
- `YamlLoadTreeBenchmarks`: load the representation tree. OfficeIMO.Yaml decodes scalar values on first use, so its tree lane leaves them undecoded; the `OfficeIMOYamlWithValues` lane also reads every scalar value as a string, which the reference trees do while loading.
- `YamlLoadAndSaveBenchmarks`: load and save without edits. OfficeIMO.Yaml saves an unchanged document by copying its retained source, and setup checks that its output equals the input.
- `YamlEmitBenchmarks`: emit a new document from events captured before timing.

Every class runs these workloads:

| Workload | Input |
| --- | --- |
| `SmallConfig` | the corpus GitHub Actions workflow |
| `KubernetesStream` | a generated multi-document Deployment and Service stream of about 10 MB |
| `DeepNesting` | 300 generated documents nesting 57 collections, below SharpYaml's 64-level parser default |
| `BlockScalars` | about 2 MB of generated literal and folded block scalars |
| `FlowHeavy` | about 2 MB of generated JSON parsed as YAML |
| `AnchorHeavy` | the corpus Airflow Compose file with anchors, aliases and merge keys |
| `ManySmallFiles` | 1,000 small corpus files processed one by one |

Generated inputs are built at run time and never committed. Setup validates each operation before timing, with the conformance adapters as an independent view: parsed events agree with the other library, loaded values agree with the other library, saved output loads to the same values, and emitted output parses back to the same events. A validation failure stops that benchmark instead of measuring broken work. `--validate-workloads` runs the same checks without timing.

Compare rows within one class and workload. Ratios use YamlDotNet as the baseline; allocation columns come from the memory diagnoser. Runs depend on the host, so keep them as opt-in evidence rather than pull-request gates.
