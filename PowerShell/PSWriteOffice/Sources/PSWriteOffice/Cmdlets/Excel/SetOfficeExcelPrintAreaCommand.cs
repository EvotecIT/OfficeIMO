using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets the print area for a worksheet.</summary>
/// <example>
///   <summary>Set a print area and keep the workbook printable.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Set-OfficeExcelPrintArea -Path .\Report.xlsx -Sheet Summary -Range A1:H40
///     Set-OfficeExcelPrintTitles -Path .\Report.xlsx -Sheet Summary -FirstRow 1 -LastRow 2
/// )
/// $proof</code>
///   <para>Stores the worksheet-local Excel print area and repeats the report header rows for printing.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelPrintArea", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelPrintArea")]
public sealed class SetOfficeExcelPrintAreaCommand : PSCmdlet {
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

    /// <summary>A1 range to print.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Range { get; set; } = string.Empty;

    /// <summary>Emit the worksheet after setting the print area.</summary>
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
        document.SetPrintArea(sheet, Range, save: false);
        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            WriteObject(sheet);
        }
    }
}