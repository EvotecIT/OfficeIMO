using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets the worksheet that opens as the active sheet.</summary>
/// <example>
///   <summary>Open the workbook on the Summary sheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$sheet = Set-OfficeExcelActiveSheet -Path .\Report.xlsx -Sheet Summary -PassThru
/// Get-OfficeExcelWorksheetView -Path .\Report.xlsx -Sheet $sheet.Name |
///     Select-Object SheetName, View, TopLeftCell</code>
///   <para>Updates workbook view state so spreadsheet applications open on Summary.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelActiveSheet", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelActiveSheet")]
[OutputType(typeof(ExcelSheet), typeof(PSObject))]
public sealed class SetOfficeExcelActiveSheetCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name to activate. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    [Alias("WorksheetName")]
    public string? Sheet { get; set; }

    /// <summary>Zero-based worksheet index to activate.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(ParameterSetName = ParameterSetPath)]
    public int? SheetIndex { get; set; }

    /// <summary>Emit the activated worksheet.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (!string.IsNullOrWhiteSpace(Sheet) && SheetIndex.HasValue) {
            throw new PSArgumentException("Specify either -Sheet or -SheetIndex, not both.");
        }

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);

        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {

            return;

        }

        ExcelSheet sheet = ResolveTargetSheet(workbook.Document);
        workbook.Document.SetActiveWorksheet(sheet);
        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            WriteObject(string.Equals(ParameterSetName, ParameterSetPath, StringComparison.OrdinalIgnoreCase)
                ? CreatePathRecord(workbook.Document, sheet)
                : sheet);
        }
    }

    private PSObject CreatePathRecord(ExcelDocument document, ExcelSheet sheet) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("Path", document.FilePath ?? SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path)));
        item.Properties.Add(new PSNoteProperty("Name", sheet.Name));
        item.Properties.Add(new PSNoteProperty("SheetName", sheet.Name));
        item.Properties.Add(new PSNoteProperty("SheetIndex", ResolveSheetIndex(document, sheet)));
        return item;
    }

    private static int ResolveSheetIndex(ExcelDocument document, ExcelSheet sheet) {
        for (int i = 0; i < document.Sheets.Count; i++) {
            if (string.Equals(document.Sheets[i].Name, sheet.Name, StringComparison.OrdinalIgnoreCase)) {
                return i;
            }
        }

        return -1;
    }

    private ExcelSheet ResolveTargetSheet(ExcelDocument document) {
        if (ParameterSetName == ParameterSetContext) {
            var context = ExcelDslContext.Require(this);
            return context.RequireSheet();
        }

        if (!string.IsNullOrWhiteSpace(Sheet)) {
            return ExcelSheetResolver.Resolve(document, Sheet, null);
        }

        if (SheetIndex.HasValue) {
            return ExcelSheetResolver.Resolve(document, null, SheetIndex);
        }

        throw new PSArgumentException("Specify -Sheet or -SheetIndex.");
    }
}