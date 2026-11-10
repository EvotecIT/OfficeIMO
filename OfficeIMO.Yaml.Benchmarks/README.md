# OfficeIMO YAML Benchmarks

This opt-in project measures OfficeIMO.Yaml and the reference YAML libraries, YamlDotNet 18.0.0, SharpYaml 3.17.0 and VYaml 1.6.1, against the same conformance data and workloads. The libraries are referenced here only; OfficeIMO runtime projects do not depend on this project.

## Conformance comparison

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Yaml.Benchmarks -- --conformance --output yaml-conformance.md
```

The command runs each library through the harness in `OfficeIMO.Yaml.Tests/Conformance`, which this project compiles by link, and prints a Markdown report. OfficeIMO.Yaml runs twice: with `Strict`, which the specification's error cases measure, and in its default mode, which also accepts the de-facto indentation listed in the reader conformance tests. Add `--details` to list every failing case and file, and `--timeout <seconds>` to change the per-call limit (default 10).

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

The Event content column compares only kinds, values, tags, anchors and alias names. It is the only event column for VYaml, whose parser does not expose scalar styles, flow markers or explicit document markers. VYaml's JSON column uses its dynamic deserialization and therefore its own scalar resolution, because its parser cannot tell a quoted scalar from a plain one. VYaml has no document model, so it has no round-trip rows, and its adapter raises the parser depth limit from 64 to OfficeIMO.Yaml's 256.

Each adapter decodes bytes with YAML byte order mark and null-pattern detection and re-encodes saved text the same way, because both libraries read text rather than bytes. JSON loading uses each library's tree loader with one shared YAML 1.2 core-schema scalar resolver, so the JSON column compares structure loading rather than resolver choices. The SharpYaml syntax tree row reports only round-trips: the library documents it as reproducing its source text, and it has no event or value view.

## Throughput workloads

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Yaml.Benchmarks -- --validate-workloads
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Yaml.Benchmarks -- --filter "*YamlParseEventsBenchmarks*"
```

Four benchmark classes cover the document operations, each with YamlDotNet as the baseline:

- `YamlParseEventsBenchmarks`: parse to events without building a tree. The OfficeIMO.Yaml lane reads every scalar value, alias name, anchor and tag as a string, matching the reference parsers' events, which carry them.
- `YamlLoadTreeBenchmarks`: load the representation tree. OfficeIMO.Yaml decodes scalar values on first use, so its tree lane leaves them undecoded; the `OfficeIMOYamlWithValues` lane also reads every scalar value as a string, which the reference trees do while loading.
- `YamlLoadAndSaveBenchmarks`: load and save without edits. OfficeIMO.Yaml saves an unchanged document by copying its retained source, and setup checks that its output equals the input.
- `YamlEmitBenchmarks`: emit a new document from events captured before timing.

VYaml runs in the parse and load classes only. It reads UTF-8, so each input is encoded once during setup. Its load lane is dynamic deserialization into dictionaries, lists and resolved scalars, because VYaml has no representation tree; setup checks those values against YamlDotNet. It has no load-and-save lane because it has no document model, and no emit lane because its emitter writes values rather than replaying parser events with their anchors and tags.

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

## Serializer

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Yaml.Benchmarks -- --filter "*YamlSerializeBenchmarks*" "*YamlDeserializeBenchmarks*" --job short
```

`YamlSerializeBenchmarks` and `YamlDeserializeBenchmarks` convert objects to and from YAML with each library, with YamlDotNet's reflection serializer as the baseline. OfficeIMO.Yaml runs twice: with its generated context and with reflection metadata. SharpYaml uses its generated context, and VYaml its generated formatters. All lanes share one model set with camelCase keys. VYaml writes and reads UTF-8 bytes, and the other libraries use strings.

| Workload | Input |
| --- | --- |
| `SmallConfig` | one configuration object with scalars, a list, a dictionary and a nested object |
| `LargeRecords` | 10,000 flat order records |
| `KubernetesNested` | 200 Deployment-shaped objects nesting metadata, label dictionaries, containers, ports, environment and resources |

Setup checks every lane before timing:

- each library's output reads back with OfficeIMO.Yaml into the workload's objects;
- each library reads OfficeIMO.Yaml's output into the same objects;
- objects are compared through OfficeIMO.Yaml's deterministic output.

`--validate-workloads` runs the same checks.

## Merge

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Yaml.Benchmarks -- --filter "*YamlMergeBenchmarks*" --job short
```

`YamlMergeBenchmarks` merges a staging overlay into an umbrella chart's values file and saves the result. The base nests the committed Bitnami nginx values under eight subchart keys (about 400 KB) and the overlay repeats the committed staging overlay under each. The YamlDotNet baseline loads both files into `YamlStream` trees, merges them by Helm's rules (maps recursively, everything else replaces) and serializes the result; the OfficeIMO.Yaml lane loads both, calls `YamlDocument.Merge` and saves. Setup checks that both outputs load in YamlDotNet to the data an independent Helm-rules merge computes. The OfficeIMO.Yaml output is also a minimal diff of the base, which YamlDotNet's serializer does not attempt.

## Large real-world corpus

The committed corpus is small enough for every test run. A larger corpus of about 13,000 files from CI systems, Kubernetes and Helm, Compose, OpenAPI, Ansible, CloudFormation and SAM, Home Assistant, documentation front matter and application configuration is pinned in [`sources.json`](../OfficeIMO.TestAssets/Yaml/LargeCorpus/sources.json) and fetched into a local cache:

```sh
pwsh ./OfficeIMO.TestAssets/Yaml/LargeCorpus/Update-YamlLargeCorpus.ps1
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Yaml.Benchmarks -- --large-corpus
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Yaml.Benchmarks -- --mutations --seed 20261110
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Yaml.Benchmarks -- --merge-corpus
```

The fetcher checks out only the selected files of each source at its pinned commit, records each file's size and SHA-256 in `index.json` and records templated files (Helm Go templates, conda Jinja recipes), Markdown without YAML front matter and out-of-bounds files as excluded with a reason. The cache is `OFFICEIMO_YAML_CORPUS_CACHE` or `yaml-corpus` under the temporary folder; `--cache` overrides it for both runs, and nothing from it is committed.

`--large-corpus` parses, loads and round-trips every indexed file with each library. It classifies each file by acceptance: a file OfficeIMO.Yaml rejects is invalid when YamlDotNet and SharpYaml also reject it, and a file to investigate when either accepts it. It compares events and loaded YAML 1.2 core values with each reference, counts how many files a YAML 1.1 loader would read differently, compares the default mode with `Strict` (files only the default mode parses, its indentation warnings and merge-key notes, and whether each merge key names a mapping), and totals throughput. It writes `large-corpus.md` and `large-corpus-failures.tsv` to `reports` in the cache, or to `--output-dir`.

`--merge-corpus` merges a synthetic overlay into every Helm `values*.yaml` file of the cache and runs the committed merge pairs' checks on each: the diff touches only reported paths, YamlDotNet loads the result to the Helm-rules reference merge, SharpYaml and OfficeIMO.Yaml load it alike, and a second merge writes the same bytes. Each overlay is derived from its base through YamlDotNet: every scalar leaf takes a new value of the same type, every mapping gets one new key, and sequences are restated with their changed items. It writes `merge-corpus.md` and `merge-corpus-failures.tsv`; `--limit` caps the file count.

`--mutations` truncates, flips bytes, edits lines and indentation, injects deep nesting, alias bombs, huge scalars, node floods, invalid UTF-8, control characters and encoding changes in a seeded sample of corpus files. It runs each mutant through the reader and the model with a per-stage timeout. A mutant passes when only `YamlException`, `YamlLimitExceededException` or `OperationCanceledException` escapes, nothing hangs and allocation stays within budget. Mutants depend only on the seed, file, kind and variant, so `--file` and `--kind` reproduce one exactly. Failing mutants are saved next to `mutations.md`.
