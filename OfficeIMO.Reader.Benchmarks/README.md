# OfficeIMO.Reader.Benchmarks

`OfficeIMO.Reader.Benchmarks` provides repeatable performance and allocation evidence for the public Reader surfaces. It is a benchmark-only project and is not packaged with the runtime libraries.

## Coverage

- rich `ReadDocument(...)` extraction across CSV, EPUB, Excel, HTML, JSON, Markdown, PDF, PowerPoint, RTF, Visio, Word, XML, YAML, and ZIP
- bounded content detection for text, signature-based, and ZIP-container formats
- version 5 JSON transport serialization and deserialization
- token-aware hierarchy construction and hierarchy JSON serialization
- Markdown parser, heading/table chunking, and paragraph-only chunking isolation for regression diagnosis

The corpus is generated deterministically during benchmark setup. Document creation is outside measured operations, while every measured read starts from the same immutable byte payload. Hashing is disabled in the extraction lane so format parsing and result projection remain visible; hosts that rely on source hashing should benchmark that option separately for their storage layer.

## Compare semantic extraction

The `compare` command writes a format-neutral evidence corpus for DOCX, XLSX, PPTX, PDF, HTML, CSV, MSG, EPUB, ZIP, and malformed input. It scores Markdown retention separately from OfficeIMO-native tables, links, assets, and source locations, then records repeatability hashes and diagnostic runtime/allocation measurements:

```powershell
dotnet run --project OfficeIMO.Reader.Benchmarks/OfficeIMO.Reader.Benchmarks.csproj -c Release -f net8.0 -- compare --output artifacts/reader-comparison
```

Competitor runs are optional. Configure direct executable invocations with `Comparison/runners.example.json`; the harness does not invoke a shell, install tools, or add them to the OfficeIMO dependency graph:

```powershell
dotnet run --project OfficeIMO.Reader.Benchmarks/OfficeIMO.Reader.Benchmarks.csproj -c Release -f net8.0 -- compare `
  --output artifacts/reader-comparison `
  --runners OfficeIMO.Reader.Benchmarks/Comparison/runners.example.json
```

The example commands assume the selected tools are already available on `PATH`. Unavailable runners are reported as evidence gaps. Generated corpus files, extracted Markdown, and machine-specific reports belong under `artifacts/` and should not be committed. Runtime values from this command include tool startup and are diagnostic; use the BenchmarkDotNet lanes below for release performance claims.

## Run

Run the complete short benchmark suite from the repository root:

```powershell
dotnet run --project OfficeIMO.Reader.Benchmarks/OfficeIMO.Reader.Benchmarks.csproj -c Release -f net8.0
```

Run one lane or format:

```powershell
dotnet run --project OfficeIMO.Reader.Benchmarks/OfficeIMO.Reader.Benchmarks.csproj -c Release -f net8.0 -- --filter *ReaderDetectionBenchmarks*
dotnet run --project OfficeIMO.Reader.Benchmarks/OfficeIMO.Reader.Benchmarks.csproj -c Release -f net8.0 -- --filter *ReaderHierarchicalChunkingBenchmarks*
dotnet run --project OfficeIMO.Reader.Benchmarks/OfficeIMO.Reader.Benchmarks.csproj -c Release -f net8.0 -- --filter "*ReaderDocumentBenchmarks*Pdf*"
```

Use the in-process dry job only to smoke-test benchmark discovery and corpus validity. Its timings are not release evidence:

```powershell
dotnet run --project OfficeIMO.Reader.Benchmarks/OfficeIMO.Reader.Benchmarks.csproj -c Release -f net8.0 -- --job Dry
```

BenchmarkDotNet writes detailed Markdown, CSV, and JSON results beneath `BenchmarkDotNet.Artifacts`. Keep large or machine-specific runs out of the repository; publish only a concise environment-qualified summary when a release decision needs a durable baseline.
