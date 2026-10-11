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

For your own CSV, pass `-InputPath`. Set the review date with `-AsOfDate yyyy-MM-dd`. The budget uses USD with at most two decimal places; positive forecast variance means overspend. Milestone dates use `yyyy-MM-dd`, progress is an integer from 0 to 100, and an incomplete milestone is overdue only when its due date precedes the review date.

If downloading a script separately, save it as a `.ps1` file and pass its companion CSV with `-InputPath`, or keep the CSV under the script's `Data` directory. The checked-in data is synthetic and includes Unicode owner names and a zero-budget spending exception.

The project CSV stores whole progress values from 0 to 100; the workbook stores percentage cells, so enter `75%` for seventy-five percent. The workbook formulas recalculate when their input cells change. Word and PDF files remain dated snapshots. Update the CSV and regenerate those copies after changes; editing a workbook does not refresh an existing report. The sample reviews fit on one page. Larger inputs can span multiple pages, so inspect pagination and fonts before distribution.
