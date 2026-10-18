#pragma warning disable CS1591
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Runs OfficeIMO safe workbook repairs for common package, table, view, print, drawing, and calculation artifacts.</summary>
/// <example>
///   <summary>Repair workbook package artifacts before handing the file to users.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$repair = Repair-OfficeExcelWorkbook -Path .\QuarterlyReport.xlsx -PassThru
/// $repair.Actions | Format-Table Category,SheetName,Message
/// if ($repair.After.HasErrors) {
///     $repair.After.Issues | Format-Table Severity,Category,SheetName,Address,Message
/// }</code>
///   <para>Uses the reusable OfficeIMO repair pipeline. The command normalizes safe workbook artifacts and returns before/after diagnostics when -PassThru is used.</para>
/// </example>
[Cmdlet(VerbsDiagnostic.Repair, "OfficeExcelWorkbook", DefaultParameterSetName = ParameterSetPath, SupportsShouldProcess = true)]
[Alias("ExcelWorkbookRepair", "ExcelRepair")]
[OutputType(typeof(PSObject))]
public sealed class RepairOfficeExcelWorkbookCommand : PSCmdlet {
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook path to repair.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open workbook document to repair.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Skip defined-name repairs.</summary>
    [Parameter]
    public SwitchParameter SkipDefinedNames { get; set; }

    /// <summary>Skip worksheet table repairs.</summary>
    [Parameter]
    public SwitchParameter SkipTables { get; set; }

    /// <summary>Skip worksheet view and freeze-pane repairs.</summary>
    [Parameter]
    public SwitchParameter SkipSheetViews { get; set; }

    /// <summary>Skip print, page-break, and page-scale repairs.</summary>
    [Parameter]
    public SwitchParameter SkipPrintSettings { get; set; }

    /// <summary>Skip drawing, image, and header/footer picture repairs.</summary>
    [Parameter]
    public SwitchParameter SkipDrawings { get; set; }

    /// <summary>Skip calculation-chain cleanup and recalc-on-open metadata.</summary>
    [Parameter]
    public SwitchParameter SkipCalculation { get; set; }

    /// <summary>Do not save after applying repairs to an open document.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

    /// <summary>Emit a repair report object.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    protected override void ProcessRecord() {
        var shouldProcessChecked = false;
        if (ParameterSetName == ParameterSetPath) {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
            if (!ShouldProcess(resolvedPath, "Repair Excel workbook")) {
                return;
            }

            shouldProcessChecked = true;
        }

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        if (!shouldProcessChecked &&
            !ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Repair Excel workbook")) {
            return;
        }

        var report = workbook.Document.RepairWorkbook(new ExcelWorkbookRepairOptions {
            DefinedNames = !SkipDefinedNames.IsPresent,
            Tables = !SkipTables.IsPresent,
            SheetViews = !SkipSheetViews.IsPresent,
            PrintSettings = !SkipPrintSettings.IsPresent,
            Drawings = !SkipDrawings.IsPresent,
            Calculation = !SkipCalculation.IsPresent,
            Save = !NoSave.IsPresent
        });

        if (!NoSave.IsPresent) {
            workbook.SaveIfOwned();
        }

        if (PassThru.IsPresent) {
            var output = new PSObject();
            output.Properties.Add(new PSNoteProperty("Path", workbook.Document.FilePath));
            output.Properties.Add(new PSNoteProperty("ActionCount", report.ActionCount));
            output.Properties.Add(new PSNoteProperty("Actions", report.Actions.Select(CreateAction).ToArray()));
            output.Properties.Add(new PSNoteProperty("Before", report.Before));
            output.Properties.Add(new PSNoteProperty("After", report.After));
            WriteObject(output);
        }
    }

    private static PSObject CreateAction(ExcelWorkbookRepairAction action) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("Category", action.Category));
        item.Properties.Add(new PSNoteProperty("SheetName", action.SheetName));
        item.Properties.Add(new PSNoteProperty("Message", action.Message));
        return item;
    }
}
#pragma warning restore CS1591