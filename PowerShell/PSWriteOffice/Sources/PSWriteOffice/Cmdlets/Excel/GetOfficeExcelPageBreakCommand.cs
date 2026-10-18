using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets manual row and column page breaks from Excel worksheets.</summary>
/// <example>
///   <summary>List manual print page breaks.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$breaks = Get-OfficeExcelPageBreak -Path .\Report.xlsx -Sheet Data
/// $breaks |
///     Sort-Object Type, Position</code>
///   <para>Returns row and column page-break records for print-layout audits.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelPageBreak", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelPageBreaks")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelPageBreakCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to inspect.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to inspect outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name to inspect. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) to inspect. Defaults to the current DSL sheet or all workbook sheets.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>Only return row page breaks.</summary>
    [Parameter]
    public SwitchParameter Row { get; set; }

    /// <summary>Only return column page breaks.</summary>
    [Parameter]
    public SwitchParameter Column { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var path = string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase)
            ? Path
            : null;
        bool includeRows = Row.IsPresent || !Column.IsPresent;
        bool includeColumns = Column.IsPresent || !Row.IsPresent;

        foreach (var sheet in ExcelWorkbookCommandService.ResolveSheets(this, workbook.Document, ParameterSetName, Sheet, SheetIndex)) {
            if (includeRows) {
                foreach (var row in sheet.GetManualRowPageBreaks()) {
                    WriteObject(ExcelPageBreakRecordService.Create("Row", row, sheet.Name, path));
                }
            }

            if (includeColumns) {
                foreach (var column in sheet.GetManualColumnPageBreaks()) {
                    WriteObject(ExcelPageBreakRecordService.Create("Column", column, sheet.Name, path));
                }
            }
        }
    }
}