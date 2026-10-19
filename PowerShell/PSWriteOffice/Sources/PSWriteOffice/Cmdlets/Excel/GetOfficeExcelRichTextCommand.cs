using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Gets mixed-format rich text runs from an Excel cell.</summary>
/// <example>
///   <summary>Inspect rich text runs from a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$runs = Get-OfficeExcelRichText -Path .\Report.xlsx -Sheet Summary -Address A1
/// $runs |
///     Format-Table Index, Text, Bold, Italic, Color</code>
///   <para>Returns one object per inline rich text run with text and style properties.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeExcelRichText", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelRichTextRuns")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeExcelRichTextCommand : PSCmdlet {
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

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: true);
        var sheet = ExcelWorkbookCommandService.ResolveSheet(this, workbook.Document, ParameterSetName, Sheet, SheetIndex);
        var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);
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