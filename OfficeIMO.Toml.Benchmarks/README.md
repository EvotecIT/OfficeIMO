# OfficeIMO TOML Benchmarks

This opt-in project measures OfficeIMO.Toml against the reference .NET TOML libraries, [Tomlyn](https://github.com/xoofx/Tomlyn) 2.10.1 and [Tomlet](https://github.com/SamboyCoding/Tomlet) 6.2.0, on the same inputs. The libraries are referenced here only; OfficeIMO runtime projects do not depend on this project, and it is not part of the solution.

## Running

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Toml.Benchmarks -- --validate-workloads
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Toml.Benchmarks -- --filter "*TomlParseBenchmarks*"
```

`--validate-workloads` checks, for every workload, that all libraries see the same number of values when parsing and loading, that OfficeIMO and Tomlyn's syntax tree reproduce each input exactly when saved, and that the serializer lanes read and write the same models. Every benchmark class runs the same check in its setup, so no lane can win by reading less.

## Operations

- `TomlParseBenchmarks`: one pass over the text without a model. OfficeIMO's `TomlReader` against Tomlyn's incremental `TomlParser`; both lanes read every key and string value as a string. Tomlet has no pull parser.
- `TomlLoadBenchmarks`: load the document model. Tomlyn deserializes into its `TomlTable` model and Tomlet parses its `TomlDocument`. OfficeIMO decodes values on first read, so `OfficeIMOToml` leaves them undecoded and `OfficeIMOTomlWithValues` also decodes every value, as the reference loaders do while loading.
- `TomlLoadAndSaveBenchmarks`: load and save without edits. OfficeIMO and Tomlyn's full-fidelity syntax tree return the input; Tomlet re-serializes its model, which is not lossless, and its saver rejects a valid key in the corpus, so that lane reports NA there.
- `TomlDeserializeBenchmarks` and `TomlSerializeBenchmarks`: map a corpus file to and from its model (`PyProject`, `CargoWorkspaceManifest`, `UvLock`) with each library's serializer, through reflection and through generated metadata (`OfficeIMOBenchmarkContext`, `TomlynBenchmarkContext`). Both libraries use a kebab-case naming policy and skip undeclared keys. Setup checks that every lane reads the same members and values, that Tomlyn reads OfficeIMO's output and OfficeIMO reads Tomlyn's to the same model, so no lane wins by reading or writing less. Each library writes the model it read, because untyped members hold its own model types. These lanes compare OfficeIMO with Tomlyn only.

## Workloads

| Workload | Input |
| --- | --- |
| `SmallConfig` | PyPA's commented sample `pyproject.toml` (7 KB) |
| `CargoWorkspace` | uv's workspace `Cargo.toml`, hundreds of inline-table dependencies (25 KB) |
| `LockFile` | uv's `uv.lock`, thousands of `[[package]]` tables and inline tables (350 KB) |
| `Corpus` | every file of the [real-world corpus](../OfficeIMO.TestAssets/Toml/Corpus/manifest.json) except uv's `ruff.toml`, whose TOML 1.1 multi-line inline table Tomlet cannot read |

## Large real-world corpus

About 5,400 real-world files (crates.io and workspace manifests, `Cargo.lock`, `pyproject.toml`, `uv.lock`, `poetry.lock`, editor configuration and Hugo sites) are pinned by source URL and SHA-256 in [`sources.json`](../OfficeIMO.TestAssets/Toml/LargeCorpus/sources.json) and fetched into a local cache that is never committed:

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Toml.Benchmarks -- --fetch-corpus
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Toml.Benchmarks -- --large-corpus
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Toml.Benchmarks -- --mutations --seed 20261113
```

`--large-corpus` loads every file, compares its data with Tomlyn's, saves it back, reads it through the stream reader and applies a scripted edit whose saved text must reload to the edited model with a minimal diff. It exits non-zero only for OfficeIMO failures (crash, round-trip, stream reader, edit); rejected files and differences from Tomlyn are listed for review. `--mutations` runs hostile variants of a seeded sample through the loader and the stream reader with time and allocation limits. The [TOML engine design](../Docs/officeimo.toml-engine-design.md#running-the-large-corpus) describes the options and the last results.

Measurements are evidence for a specific machine and runtime. Keep them out of pull-request gates, as the repository's benchmark rules require.
