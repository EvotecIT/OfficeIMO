using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets mixed-format rich text runs in an Excel cell.</summary>
/// <example>
///   <summary>Write a status label with a bold colored value.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Summary' {
///     Set-OfficeExcelRichText -Address A1 -Run 'Status: ', @{ Text = 'Blocked'; Bold = $true; Color = '#C00000' }
/// }</code>
///   <para>Stores inline rich text in the target cell using OfficeIMO's reusable rich-text cell model.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelRichText", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelRichText")]
public sealed class SetOfficeExcelRichTextCommand : PSCmdlet {
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

    /// <summary>1-based row index.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    public int? Row { get; set; }

    /// <summary>1-based column index.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    public int? Column { get; set; }

    /// <summary>A1-style cell address.</summary>
    [Parameter(ParameterSetName = ParameterSetContext)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? Address { get; set; }

    /// <summary>Rich text runs. Each run can be a string, hashtable, PSCustomObject, or ExcelRichTextRun.</summary>
    [Parameter(Mandatory = true)]
    [Alias("Runs")]
    public object[] Run { get; set; } = [];

    /// <summary>Emit written rich text runs.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var runs = ExcelRichTextRunService.ToRuns(Run);
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {
            return;
        }

        var sheet = ExcelWorkbookCommandService.ResolveSheet(this, workbook.Document, ParameterSetName, Sheet, SheetIndex);
        var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);
        sheet.SetRichText(row, column, runs);
        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            WriteRuns(sheet, row, column);
        }
    }

    private void WriteRuns(ExcelSheet sheet, int row, int column) {
        var address = A1.CellReference(row, column);
        var path = string.Equals(ParameterSetName, ParameterSetPath, System.StringComparison.OrdinalIgnoreCase)
            ? Path
            : null;
        var runs = sheet.GetRichText(row, column);
        for (var index = 0; index < runs.Count; index++) {
            WriteObject(ExcelRichTextRunService.CreateRecord(runs[index], index, address, row, column, sheet.Name, path));
        }
    }
}