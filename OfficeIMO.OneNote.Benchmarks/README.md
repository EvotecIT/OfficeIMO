# OfficeIMO.OneNote benchmarks

This permanent BenchmarkDotNet suite tracks how native desktop `.one` reading, writing, and semantic Markdown projection scale from one page to a representative multi-page section. The useful comparison is input scale and later before/after runs; exact timings vary by machine.

Validate the suite quickly:

```powershell
dotnet run -c Release --framework net8.0 -- --filter "*OneNoteReadWriteBenchmarks*" --job Dry --noOverwrite
```

For measurements, use `--job Short` while iterating and the default job for recorded release evidence. Keep benchmark artifacts outside the repository or delete them after summarizing the result.

## Workflow evidence

The opt-in evidence runner measures complete section create/write, read, and
read/edit/write workflows at 1-page, 25-page, and 100-page scales. Every sample
runs in a separate child process and records elapsed time, managed allocations,
peak working set, input size, output size, page and paragraph counts, and a
SHA-256 structural fingerprint over ordered section, page, paragraph, run text,
and run-style fields. Validation runs outside the timed and allocation window.
File-producing lanes are reopened after measurement and must exactly match the
expected semantic fingerprint.

Run the representative 25-page matrix three times and keep the JSON under the
ignored benchmark artifact root:

```powershell
dotnet run -c Release --framework net10.0 -- --evidence --scale Normal --repeat 3 --json .benchmark-artifacts/onenote/normal.json
```

Use `--operation CreateWrite|Read|ReadEditWrite` or
`--scale Small|Normal|Large` to narrow an investigation. The report records the
source commit, dirty-tree state, runtime, operating system, architecture, and
logical processor count. Compare peak working set only between isolated probes
on similar environments; use repeated samples before defining a regression
budget.
