using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds manual row or column page breaks to an Excel worksheet.</summary>
/// <example>
///   <summary>Insert print page breaks for report sections.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelPageBreak -Row 25,50 -Column 8 }</code>
///   <para>Adds row page breaks after rows 25 and 50 and a column page break after column 8.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelPageBreak", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelPageBreak")]
public sealed class AddOfficeExcelPageBreakCommand : PSCmdlet {
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

    /// <summary>Worksheet name. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>One-based rows after which manual page breaks should be inserted.</summary>
    [Parameter]
    public int[] Row { get; set; } = [];

    /// <summary>One-based columns after which manual page breaks should be inserted.</summary>
    [Parameter]
    public int[] Column { get; set; } = [];

    /// <summary>Emit page-break records after adding them.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Row.Length == 0 && Column.Length == 0) {
            throw new PSArgumentException("Provide at least one row or column page break.");
        }

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);

        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {

            return;

        }

        var sheet = ExcelWorkbookCommandService.ResolveSheet(this, workbook.Document, ParameterSetName, Sheet, SheetIndex);

        foreach (var row in Row) {
            sheet.AddManualRowPageBreak(row, save: false);
        }

        foreach (var column in Column) {
            sheet.AddManualColumnPageBreak(column, save: false);
        }

        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            WritePageBreaks(sheet);
        }
    }

    private void WritePageBreaks(ExcelSheet sheet) {
        foreach (var row in sheet.GetManualRowPageBreaks()) {
            WriteObject(ExcelPageBreakRecordService.Create("Row", row, sheet.Name, PathForRecord()));
        }

        foreach (var column in sheet.GetManualColumnPageBreaks()) {
            WriteObject(ExcelPageBreakRecordService.Create("Column", column, sheet.Name, PathForRecord()));
        }
    }

    private string? PathForRecord() {
        return string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase)
            ? Path
            : null;
    }
}