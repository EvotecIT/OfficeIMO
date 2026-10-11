# PowerShell document workflows

These scripts use PSWriteOffice to create editable documents and PDF copies from the same input. Install the public module version selected by the command, then run from the OfficeIMO checkout:

```powershell
Install-Module PSWriteOffice -RequiredVersion 3.0.8 -Scope CurrentUser
pwsh -NoProfile -File OfficeIMO.Examples/PowerShell/BudgetForecastReview.ps1 -OutputDirectory ./out -ModuleVersion 3.0.8
pwsh -NoProfile -File OfficeIMO.Examples/PowerShell/ProjectDeliveryPack.ps1 -OutputDirectory ./out -ModuleVersion 3.0.8
```

Use PowerShell 7.4 or newer, or replace `pwsh` with `powershell` for Windows PowerShell 5.1 on Windows. The scripts import the requested module version explicitly.

| Workflow | Input | Output and decision |
| --- | --- | --- |
| `PatchComplianceReport.ps1` | Synthetic site inventory in the script | DOCX and PDF with compliance, exceptions, chart and maintenance actions |
| `BudgetForecastReview.ps1` | `Data/BudgetForecast.csv` | XLSX detail and formulas, plus a DOCX/PDF review naming overspend decisions |
| `ProjectDeliveryPack.ps1` | `Data/ProjectMilestones.csv` | XLSX milestone register and status formulas, plus a DOCX/PDF review naming overdue work and recovery actions |

For your own CSV, pass `-InputPath`. Set the review date with `-AsOfDate yyyy-MM-dd`. The budget uses non-negative USD amounts with a dot as the decimal separator, at most two decimal places, and no grouping separators: use `1234.50`. Positive forecast variance means overspend. Milestone dates use `yyyy-MM-dd`, progress is an integer from 0 to 100, and an incomplete milestone is overdue only when its due date precedes the review date.

If downloading a script separately, save it as a `.ps1` file and pass its companion CSV with `-InputPath`, or keep the CSV under the script's `Data` directory. The checked-in data is synthetic and includes Unicode owner names and a zero-budget spending exception. Reports generated from the default input carry a synthetic-sample label. When you pass `-InputPath`, the report uses a neutral snapshot label.

The project CSV stores whole progress values from 0 to 100; the workbook stores percentage cells, so enter `75%` for seventy-five percent. The workbook formulas recalculate when their input cells change. Word and PDF files remain dated snapshots. Update the CSV and regenerate those copies after changes; editing a workbook does not refresh an existing report. The sample reviews fit on one page. Larger inputs can span multiple pages, so inspect pagination and fonts before distribution.

To refresh the website downloads, use the existing evidence builder:

```powershell
./Website/scripts/Build-ShowcaseEvidence.ps1 -ExampleId pswriteoffice-budget-review,pswriteoffice-project-delivery -DocumentsRoot ./out
```

Its `Website/static/downloads/showcase/manifest.json` records the generating script and CSV hashes, the loaded module version and complete package-file inventory, native output hashes, PDF conversion warnings, font diagnostics, and preview renderer identity. The producer receipt has a checksum over its complete captured claim; missing or edited package, conversion or renderer records invalidate it. Each rendered preview must also match its declared PDF input and current output hash. A missing font-program version or hash is recorded as unknown; an unembedded PDF font is resolved by the viewer. Package hashes identify the producer without establishing a viewer's font selection.

The recipes, sample CSV files, generation worker, evidence builder and loaded helpers use LF line endings on every checkout so these exact hashes also match on Windows. Changing the worker, builder or a loaded helper invalidates retained producer evidence and requires generation and review again.

Inspect every generated PDF page and its website preview before recording a review. In a separate invocation, run the builder with `-ManifestOnly -ReviewedExampleId pswriteoffice-budget-review,pswriteoffice-project-delivery`; generation and review cannot be combined. This binds the completed review to the exact preview set and hashes. A missing or changed review binding returns the review to pending. A manifest-only refresh preserves matching producer evidence; changed inputs, outputs, previews, or the documented module version invalidate it and require generation and review again. `-SkipGeneration` imports existing files and does not invent a producer record for them. The website pipeline prepares the documented PSWriteOffice version and runs both the receipt mutation checks and the real child-runner omitted-export regression before its documentation contracts.
