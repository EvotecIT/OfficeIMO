using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Clears manual row or column page breaks from an Excel worksheet.</summary>
/// <example>
///   <summary>Remove stale manual page breaks.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Clear-OfficeExcelPageBreak -Path .\Report.xlsx -Sheet Data -Row 25 -Column 8 -Confirm:$false
/// Get-OfficeExcelPageBreak -Path .\Report.xlsx -Sheet Data |
///     Format-Table Type, Position, SheetName</code>
///   <para>Removes the selected row and column page breaks and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsCommon.Clear, "OfficeExcelPageBreak", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Medium, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelPageBreakClear")]
public sealed class ClearOfficeExcelPageBreakCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name to update. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) to update. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>One-based rows whose manual page breaks should be removed.</summary>
    [Parameter]
    public int[] Row { get; set; } = [];

    /// <summary>One-based columns whose manual page breaks should be removed.</summary>
    [Parameter]
    public int[] Column { get; set; } = [];

    /// <summary>Clear all manual row and column page breaks from selected worksheets.</summary>
    [Parameter]
    public SwitchParameter All { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!All.IsPresent && Row.Length == 0 && Column.Length == 0) {
            throw new PSArgumentException("Provide row breaks, column breaks, or -All.");
        }

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        var shouldSave = false;
        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, workbook.Document, ParameterSetName, Sheet, SheetIndex)) {
            if (!ShouldProcess(sheet.Name, "Clear Excel manual page breaks")) {
                continue;
            }

            if (All.IsPresent) {
                shouldSave |= sheet.ClearManualPageBreaks(save: false);
                continue;
            }

            foreach (var row in Row) {
                shouldSave |= sheet.RemoveManualRowPageBreak(row, save: false);
            }

            foreach (var column in Column) {
                shouldSave |= sheet.RemoveManualColumnPageBreak(column, save: false);
            }
        }

        if (shouldSave) {
            workbook.SaveIfOwned();
            WritePassThru(workbook.Document, workbook.OwnsDocument, Path);
        }
    }
}