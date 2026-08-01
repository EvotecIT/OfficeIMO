# OfficeIMO.PowerPoint workflow baselines

This project measures complete PowerPoint workflows against a deterministic editable deck corpus. It records elapsed time, total managed allocations, process peak working set, input/output size, slide count, and shape count. It does not enforce regression budgets yet; the checked-in evidence must first establish stable results across representative machines and runtimes.

## Run the baseline

Run all four workflows at small, normal, and large scales:

```powershell
dotnet run --project OfficeIMO.PowerPoint.Benchmarks/OfficeIMO.PowerPoint.Benchmarks.csproj -c Release -f net10.0 -- --json .benchmark-artifacts/powerpoint/baseline.json
```

Use `--scale Small`, `--scale Normal`, or `--scale Large` for a shorter run. Each operation runs in a separate child process so process memory from another workflow does not leak into its peak-working-set result.

Generate the repeatable visual corpus used for structural and visual review:

```powershell
dotnet run --project OfficeIMO.PowerPoint.Benchmarks/OfficeIMO.PowerPoint.Benchmarks.csproj -c Release -f net10.0 -- --visual-corpus .benchmark-artifacts/powerpoint/visual-corpus
```

That command creates one validated PPTX, PNG and SVG renders for all nine slides, and a PDF. The corpus covers all seven semantic SmartArt layouts, shared custom geometry, a native table and chart, classic and modern threaded comments, and a custom show. Any Open XML error, missing review metadata, export warning, or PDF failure stops the run.

## Workloads

| Workflow | Measured contract |
| --- | --- |
| `CreateSave` | Create the full editable corpus and serialize a valid PPTX package. |
| `OpenEditSave` | Open the corpus, add a review marker to every tenth slide, and serialize a valid edited PPTX package. |
| `OpenImageExport` | Open the corpus and export every slide as a non-empty PNG with valid dimensions. |
| `OpenPdfExport` | Open the corpus and export a PDF with one page for every slide. |

The scales contain 3, 30, and 120 slides. Slides use editable text, vector shapes, tables, and charts. Package lanes reopen the result, check slide and shape counts, and run Open XML validation. Export lanes verify every image and PDF page count. Validation runs outside the timed interval but in the same probe, so an invalid or incomplete result cannot be reported as a fast result.

## Interpret the result

Compare like with like: the same runtime, architecture, build configuration, scale, and operation. Treat elapsed time from a busy workstation as directional. Managed allocations are usually more stable and should expose accidental buffering or repeated model construction. Peak working set includes the runtime and native rendering dependencies, so compare it only against another isolated probe with the same environment.

Do not add budgets from a single run. Capture repeatable baselines on Windows and at least one non-Windows environment, inspect variance, then set headroom that catches regressions without failing on normal machine noise.

The current evidence and interpretation are recorded in [BASELINE.md](BASELINE.md). `OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler` is an opt-in comparison project outside the normal solution. It mirrors the editable slide dimensions, styling, text/vector/table/two-series-chart mix, edit cadence, and package validation used by the OfficeIMO create/save and open/edit/save lanes; keep its third-party dependency isolated from product and routine benchmark builds.
