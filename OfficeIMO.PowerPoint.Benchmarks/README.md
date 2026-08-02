# OfficeIMO.PowerPoint workflow baselines

This project measures complete PowerPoint workflows against a deterministic editable deck corpus. It records elapsed time, total managed allocations, process peak working set, input/output size, slide count, and shape count. It does not enforce regression budgets yet; the checked-in evidence must first establish stable results across representative machines and runtimes.

## Run the baseline

Run all four workflows at small, normal, and large scales:

```powershell
dotnet run --project OfficeIMO.PowerPoint.Benchmarks/OfficeIMO.PowerPoint.Benchmarks.csproj -c Release -f net10.0 -- --json .benchmark-artifacts/powerpoint/baseline.json
```

Use `--scale Small`, `--scale Normal`, or `--scale Large` for a shorter run. Each operation runs in a separate child process so process memory from another workflow does not leak into its peak-working-set result. For open workflows, the parent creates the source fixture before starting the probe; the probe includes reading that source from disk in the measured workflow without including fixture authoring in its process-lifetime peak.

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
| `OpenImageExport` | Open the corpus and export every slide as a decodable PNG with no unexpected diagnostics and verified visible content, including representative table and chart regions. |
| `OpenPdfExport` | Open the corpus and export a parseable PDF whose per-page text and rendered pixels preserve representative slide, table, and chart content. |

The scales contain 3, 30, and 120 slides. Slides use editable text, vector shapes, tables, and charts. Package lanes reopen the result and verify the expected text, background and card fills, table values and header styling, chart categories and series values, edit markers, slide and shape counts, and Open XML validity. Export lanes decode or parse every output and verify representative visible content, including table and chart regions. Validation runs after timing and peak-working-set capture but in the same probe, so validation work does not contaminate the reported peak and an invalid or incomplete result cannot be reported as a fast result. Probes are intentionally cold; compare only like-for-like isolated runs.

## Interpret the result

Compare like with like: the same runtime, architecture, build configuration, scale, and operation. Treat elapsed time from a busy workstation as directional. Managed allocations are usually more stable and should expose accidental buffering or repeated model construction. Peak working set includes the runtime and native rendering dependencies, so compare it only against another isolated probe with the same environment.

Do not add budgets from a single run. Capture repeatable baselines on Windows and at least one non-Windows environment, inspect variance, then set headroom that catches regressions without failing on normal machine noise.

The current evidence and interpretation are recorded in [BASELINE.md](BASELINE.md). `OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler` is an opt-in comparison project outside the normal solution. It mirrors the editable slide dimensions, styling, text/vector/table/two-series-chart mix, and edit cadence used by the OfficeIMO create/save and open/edit/save lanes. Both producers compile the same semantic validator, while the third-party dependency remains isolated from product and routine benchmark builds.
