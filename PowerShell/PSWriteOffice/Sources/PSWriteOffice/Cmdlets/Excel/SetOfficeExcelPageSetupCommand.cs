using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets orientation, margins, scaling, print area and print titles, alone or from a print preset.</summary>
/// <example>
/// <summary>Print a sheet landscape on one page wide</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
///     ExcelSheet 'Orders' {
///         ExcelPageSetup -Orientation Landscape -FitToWidth 1 -FitToHeight 0
///     }
/// }</code>
/// <para>Sets landscape orientation and fits the columns to one page width.</para>
/// </example>
/// <example>
/// <summary>Prepare a long list for printing</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Inventory.xlsx {
///     ExcelSheet 'Servers' {
///         ExcelPageSetup -Preset DataTable -PrintArea A1:H420 -RepeatRows 1:1 -MarginPreset Narrow
///     }
/// }</code>
/// <para>Starts from the data-table print preset, repeats the header row on every page and limits printing to the table area with narrow margins.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelPageSetup", DefaultParameterSetName = "Default")]
[Alias("ExcelPageSetup")]
public sealed class SetOfficeExcelPageSetupCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Margin preset to apply.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelPrintLayoutPreset? Preset { get; set; }

    /// <summary>Orientation to apply.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficePageOrientation? Orientation { get; set; }

    /// <summary>Native print-margin preset applied before explicit margin overrides.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelMarginPreset? MarginPreset { get; set; }

    /// <summary>Top page margin in PDF points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MarginTop { get; set; }

    /// <summary>Bottom page margin in PDF points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MarginBottom { get; set; }

    /// <summary>Left page margin in PDF points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MarginLeft { get; set; }

    /// <summary>Right page margin in PDF points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MarginRight { get; set; }

    /// <summary>Physical distance from the page edge to the header.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MarginHeader { get; set; }

    /// <summary>Physical distance from the page edge to the footer.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MarginFooter { get; set; }

    /// <summary>Number of pages to fit horizontally.</summary>
    [Parameter(ParameterSetName = "Default")]
    public uint? FitToWidth { get; set; }

    /// <summary>Number of pages to fit vertically.</summary>
    [Parameter(ParameterSetName = "Default")]
    public uint? FitToHeight { get; set; }

    /// <summary>Manual scale percentage (10-400).</summary>
    [Parameter(ParameterSetName = "Default")]
    public uint? Scale { get; set; }

    /// <summary>Multi-page print order.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelPageOrder? PageOrder { get; set; }

    /// <summary>Optional print area in A1 notation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? PrintArea { get; set; }

    /// <summary>Inclusive print-title row span, such as 1:2.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? RepeatRows { get; set; }

    /// <summary>Inclusive print-title column span, such as A:B.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? RepeatColumns { get; set; }

    /// <summary>Clears repeated print rows and columns.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoPrintTitles { get; set; }

    /// <summary>Do not apply print-title rows from the selected preset.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoPresetPrintTitles { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var sheet = ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
        var page = sheet.PageSetup;
        if (Preset.HasValue) page.ApplyLayout(new ExcelPrintLayoutOptions { Preset = Preset.Value, SuppressPresetPrintTitles = NoPresetPrintTitles });
        if (Orientation.HasValue) page.Orientation = Orientation;
        if (MarginPreset.HasValue) page.SetMargins(MarginPreset.Value);
        if (MarginLeft.HasValue || MarginRight.HasValue || MarginTop.HasValue || MarginBottom.HasValue || MarginHeader.HasValue || MarginFooter.HasValue) {
            var margins = page.Margins;
            page.SetMargins(MarginLeft ?? margins?.Left ?? OfficeLength.Inches(0.7), MarginRight ?? margins?.Right ?? OfficeLength.Inches(0.7), MarginTop ?? margins?.Top ?? OfficeLength.Inches(0.75), MarginBottom ?? margins?.Bottom ?? OfficeLength.Inches(0.75), MarginHeader ?? margins?.Header, MarginFooter ?? margins?.Footer);
        }
        if (FitToWidth.HasValue) page.FitToWidth = FitToWidth;
        if (FitToHeight.HasValue) page.FitToHeight = FitToHeight;
        if (Scale.HasValue) page.Scale = Scale;
        if (PageOrder.HasValue) page.PageOrder = PageOrder;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(PrintArea))) page.PrintArea = PrintArea;
        if (NoPrintTitles) page.SetPrintTitles(null, null);
        else if (RepeatRows != null || RepeatColumns != null) {
            var titles = page.PrintTitles;
            var rows = RepeatRows == null ? ((int?, int?))(titles.FirstRow, titles.LastRow) : ExcelCommandTargets.Rows(RepeatRows);
            var columns = RepeatColumns == null ? ((int?, int?))(titles.FirstColumn, titles.LastColumn) : ExcelCommandTargets.Columns(RepeatColumns);
            page.SetPrintTitles(rows.Item1, rows.Item2, columns.Item1, columns.Item2);
        }
        if (PassThru) WriteObject(sheet);
    }

}
