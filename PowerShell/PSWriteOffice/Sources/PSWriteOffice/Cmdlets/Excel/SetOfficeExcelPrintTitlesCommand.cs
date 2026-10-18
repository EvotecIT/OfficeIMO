using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets or clears repeating print title rows and columns for a worksheet.</summary>
/// <example>
///   <summary>Repeat report headers on every printed page.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Set-OfficeExcelPrintTitles -Path .\Report.xlsx -Sheet Data -FirstRow 1 -LastRow 1 -FirstColumn 1 -LastColumn 1
///     Get-OfficeExcelSummary -Path .\Report.xlsx |
///         Select-Object -Property SheetCount, TableCount
/// )
/// $proof</code>
///   <para>Stores Excel print titles for the Data worksheet and then reads back workbook structure as a quick proof step.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelPrintTitles", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelPrintTitles")]
public sealed class SetOfficeExcelPrintTitlesCommand : PSCmdlet {
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

    /// <summary>First 1-based row to repeat.</summary>
    [Parameter]
    public int? FirstRow { get; set; }

    /// <summary>Last 1-based row to repeat.</summary>
    [Parameter]
    public int? LastRow { get; set; }

    /// <summary>First 1-based column to repeat.</summary>
    [Parameter]
    public int? FirstColumn { get; set; }

    /// <summary>Last 1-based column to repeat.</summary>
    [Parameter]
    public int? LastColumn { get; set; }

    /// <summary>Clear existing print titles for the worksheet.</summary>
    [Parameter]
    public SwitchParameter Clear { get; set; }

    /// <summary>Emit the worksheet after setting print titles.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!Clear.IsPresent && !HasRows() && !HasColumns()) {
            throw new PSArgumentException("Provide row titles, column titles, or -Clear.");
        }

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);

        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {

            return;

        }

        var document = workbook.Document;
        var sheet = ExcelWorkbookCommandService.ResolveSheet(this, document, ParameterSetName, Sheet, SheetIndex);
        document.SetPrintTitles(
            sheet,
            Clear.IsPresent ? null : FirstRow,
            Clear.IsPresent ? null : LastRow,
            Clear.IsPresent ? null : FirstColumn,
            Clear.IsPresent ? null : LastColumn,
            save: false);
        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            WriteObject(sheet);
        }
    }

    private bool HasRows() {
        return FirstRow.HasValue && LastRow.HasValue;
    }

    private bool HasColumns() {
        return FirstColumn.HasValue && LastColumn.HasValue;
    }

}