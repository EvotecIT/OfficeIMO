using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Moves a worksheet to a new workbook position.</summary>
/// <example>
///   <summary>Move the summary sheet to the front and verify order.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Move-OfficeExcelSheet -Path .\Report.xlsx -Sheet Summary -Index 0
///     Get-OfficeExcelSummary -Path .\Report.xlsx -IncludeSheets |
///         Select-Object -ExpandProperty Sheets |
///         Select-Object -First 3 -Property Index, Name
/// )
/// $proof</code>
///   <para>Moves Summary to the first worksheet tab and reads back the first sheets from workbook summary.</para>
/// </example>
[Cmdlet(VerbsCommon.Move, "OfficeExcelSheet", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("Set-OfficeExcelSheetOrder", "ExcelSheetOrder")]
public sealed class MoveOfficeExcelSheetCommand : PSCmdlet {
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

    /// <summary>Worksheet name to move. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index to move when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>Zero-based destination tab index.</summary>
    [Parameter(Mandatory = true)]
    [Alias("TargetIndex")]
    public int Index { get; set; }

    /// <summary>Emit the moved worksheet.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {
            return;
        }

        var document = workbook.Document;
        var sheet = ExcelWorkbookCommandService.ResolveSheet(this, document, ParameterSetName, Sheet, SheetIndex);
        document.ReorderWorksheet(sheet, Index);
        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            WriteObject(sheet);
        }
    }
}