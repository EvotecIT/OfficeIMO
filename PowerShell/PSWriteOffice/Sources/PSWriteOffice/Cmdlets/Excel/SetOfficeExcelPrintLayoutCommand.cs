using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Applies a reusable worksheet print layout preset.</summary>
/// <example>
///   <summary>Apply a report print layout.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Report' { Set-OfficeExcelPrintLayout -Preset Report -PrintArea A1:H40 }</code>
///   <para>Applies landscape orientation, narrow margins, one-page-wide scaling, and repeated header row.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelPrintLayout", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelPrintLayout")]
[OutputType(typeof(ExcelSheet), typeof(PSObject))]
public sealed class SetOfficeExcelPrintLayoutCommand : PSCmdlet {
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

    /// <summary>Print layout preset.</summary>
    [Parameter]
    public ExcelPrintLayoutPreset Preset { get; set; } = ExcelPrintLayoutPreset.Report;

    /// <summary>Optional print area in A1 notation.</summary>
    [Parameter]
    public string? PrintArea { get; set; }

    /// <summary>Optional orientation override.</summary>
    [Parameter]
    public OfficePageOrientation? Orientation { get; set; }

    /// <summary>Optional margin preset override.</summary>
    [Parameter]
    public ExcelMarginPreset? Margins { get; set; }

    /// <summary>Optional pages-wide fit override.</summary>
    [Parameter]
    public uint? FitToWidth { get; set; }

    /// <summary>Optional pages-tall fit override. Use 0 for unlimited height.</summary>
    [Parameter]
    public uint? FitToHeight { get; set; }

    /// <summary>Optional manual scale percentage override.</summary>
    [Parameter]
    public uint? Scale { get; set; }

    /// <summary>Optional multi-page print order override.</summary>
    [Parameter]
    public ExcelPageOrder? PageOrder { get; set; }

    /// <summary>Optional first 1-based repeated print-title row.</summary>
    [Parameter]
    public int? RepeatFirstRow { get; set; }

    /// <summary>Optional last 1-based repeated print-title row.</summary>
    [Parameter]
    public int? RepeatLastRow { get; set; }

    /// <summary>Optional first 1-based repeated print-title column.</summary>
    [Parameter]
    public int? RepeatFirstColumn { get; set; }

    /// <summary>Optional last 1-based repeated print-title column.</summary>
    [Parameter]
    public int? RepeatLastColumn { get; set; }

    /// <summary>Do not apply print-title rows from the selected preset.</summary>
    [Parameter]
    public SwitchParameter NoPresetPrintTitles { get; set; }

    /// <summary>Emit the worksheet after applying the layout.</summary>
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
        sheet.ApplyPrintLayout(new ExcelPrintLayoutOptions {
            Preset = Preset,
            PrintArea = PrintArea,
            Orientation = Orientation,
            Margins = Margins,
            FitToWidth = FitToWidth,
            FitToHeight = FitToHeight,
            Scale = Scale,
            PageOrder = PageOrder,
            RepeatFirstRow = RepeatFirstRow,
            RepeatLastRow = RepeatLastRow,
            RepeatFirstColumn = RepeatFirstColumn,
            RepeatLastColumn = RepeatLastColumn,
            SuppressPresetPrintTitles = NoPresetPrintTitles.IsPresent
        });

        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            WriteObject(string.Equals(ParameterSetName, ParameterSetPath, StringComparison.OrdinalIgnoreCase)
                ? CreatePathRecord(document, sheet)
                : sheet);
        }
    }

    private PSObject CreatePathRecord(ExcelDocument document, ExcelSheet sheet) {
        var item = new PSObject();
        item.Properties.Add(new PSNoteProperty("Path", document.FilePath ?? SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path)));
        item.Properties.Add(new PSNoteProperty("Name", sheet.Name));
        item.Properties.Add(new PSNoteProperty("SheetName", sheet.Name));
        item.Properties.Add(new PSNoteProperty("SheetIndex", ResolveSheetIndex(document, sheet)));
        item.Properties.Add(new PSNoteProperty("Preset", Preset));
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
}