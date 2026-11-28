# OfficeIMO.Scaling.Benchmarks

Size-scaled lanes for formats and editing APIs that have no comparison
benchmark. Each lane builds, writes or reads a document at 10, 100 and 1,000
units (paragraphs, rows, slides, shapes, attachments, sheets, ...) and reports
time and allocation per unit. A linear operation keeps a flat per-unit cost;
one that rescans or re-serializes the document per call grows with the size.

Every lane validates its last result (unit counts after a round trip), so a
lane can't get faster by doing less work.

```powershell
# Quick stopwatch sweep: per-unit time and allocation, and the growth between the last two sizes
dotnet run -c Release -f net10.0 --project OfficeIMO.Scaling.Benchmarks -- --scan
dotnet run -c Release -f net10.0 --project OfficeIMO.Scaling.Benchmarks -- --scan --filter Word. --sizes 100,1000,5000

# BenchmarkDotNet run with µs/unit and B/unit columns
dotnet run -c Release -f net10.0 --project OfficeIMO.Scaling.Benchmarks -- --filter "*FormatScaling*"
# Only some lanes
$env:OFFICEIMO_SCALING_LANES = "Word.AddComment,Excel.AddSheet"
```

Lanes are listed in `ScalingLanes.cs`. These are measurement tools only; CI
builds the project but never runs it.
