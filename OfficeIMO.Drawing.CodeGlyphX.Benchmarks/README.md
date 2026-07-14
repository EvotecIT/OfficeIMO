# CodeGlyphX SVG import benchmarks

This non-packable BenchmarkDotNet project measures `OfficeSvgDrawingReader` against SVG produced by CodeGlyphX. SVG generation happens during benchmark setup, so the timed method measures import rather than encoding or rendering.

The scenarios cover default QR, styled QR circles, Data Matrix, DataBar Expanded stacked, and Code 128 with human-readable text.

## Validate the benchmark

Always run a dry pass before collecting measurements:

```powershell
dotnet run -c Release -- --filter "*" --job Dry --noOverwrite
```

## Measure SVG import

```powershell
dotnet run -c Release -- --filter "*" --noOverwrite
```

BenchmarkDotNet writes its Markdown and CSV reports under `BenchmarkDotNet.Artifacts`. Use `--artifacts <path>` to keep task-specific results in an explicit temporary folder.

## Report drawing complexity

The deterministic complexity mode reports SVG bytes, drawing elements, shapes, searchable text runs, and unsupported features without running timed benchmarks:

```powershell
dotnet run -c Release -- --complexity
```

By default, the project uses the published CodeGlyphX package. Pass `-p:CodeGlyphXProjectPath="<path-to-CodeGlyphX.csproj>"` before `--` to benchmark an adjacent source checkout.
