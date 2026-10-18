using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates discoverable Excel-to-PDF conversion options for Export-OfficeDocumentPdf.</summary>
/// <example>
///   <summary>Export selected visible sheets with workbook layout features.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeExcelPdfOptions -SheetName Summary,Services -UseWorksheetCharts -UseWorksheetImages
/// Export-OfficeDocumentPdf -InputPath .\Report.xlsx -Path .\Report.pdf -ExcelOptions $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcelPdfOptions")]
[OutputType(typeof(ExcelToPdfOptions))]
public sealed class NewOfficeExcelPdfOptionsCommand : PSCmdlet {
    /// <summary>Underlying low-level OfficeIMO PDF options.</summary>
    [Parameter]
    public OfficeIMO.Pdf.PdfOptions? PdfOptions { get; set; }

    /// <summary>Default font family used when the workbook does not specify one.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>PDF page size.</summary>
    [Parameter]
    public PageSize? PageSize { get; set; }

    /// <summary>Left page margin in PDF points.</summary>
    [Parameter]
    [ValidateRange(0d, double.MaxValue)]
    public double? MarginLeft { get; set; }

    /// <summary>Top page margin in PDF points.</summary>
    [Parameter]
    [ValidateRange(0d, double.MaxValue)]
    public double? MarginTop { get; set; }

    /// <summary>Right page margin in PDF points.</summary>
    [Parameter]
    [ValidateRange(0d, double.MaxValue)]
    public double? MarginRight { get; set; }

    /// <summary>Bottom page margin in PDF points.</summary>
    [Parameter]
    [ValidateRange(0d, double.MaxValue)]
    public double? MarginBottom { get; set; }

    /// <summary>Controls how worksheet content is laid out on PDF pages.</summary>
    [Parameter]
    public ExcelPdfWorksheetLayoutMode? WorksheetLayout { get; set; }

    /// <summary>Worksheet names to export. The default exports all eligible sheets.</summary>
    [Parameter]
    [Alias("SheetNames")]
    public string[]? SheetName { get; set; }

    /// <summary>Exclude workbook sheets marked hidden.</summary>
    [Parameter]
    public SwitchParameter RespectWorkbookSheetVisibility { get; set; }

    /// <summary>Honor worksheet print areas.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetPrintAreas { get; set; }

    /// <summary>Honor worksheet page setup.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetPageSetup { get; set; }

    /// <summary>Honor worksheet rows configured to repeat on printed pages.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetPrintTitleRows { get; set; }

    /// <summary>Honor worksheet page breaks.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetPageBreaks { get; set; }

    /// <summary>Render worksheet headers and footers.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetHeadersAndFooters { get; set; }

    /// <summary>Render images referenced by worksheet headers and footers.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetHeaderFooterImages { get; set; }

    /// <summary>Render worksheet cell styles.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetCellStyles { get; set; }

    /// <summary>Render worksheet hyperlinks.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetHyperlinks { get; set; }

    /// <summary>Render worksheet images.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetImages { get; set; }

    /// <summary>Render worksheet charts.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetCharts { get; set; }

    /// <summary>Chart visual style override.</summary>
    [Parameter]
    public OfficeChartStyle? ChartStyle { get; set; }

    /// <summary>Chart layout override.</summary>
    [Parameter]
    public OfficeChartLayout? ChartLayout { get; set; }

    /// <summary>Render merged worksheet cells.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetMergedCells { get; set; }

    /// <summary>Honor worksheet column widths.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetColumnWidths { get; set; }

    /// <summary>Honor worksheet row heights.</summary>
    [Parameter]
    public SwitchParameter UseWorksheetRowHeights { get; set; }

    /// <summary>Exclude hidden worksheet rows and columns.</summary>
    [Parameter]
    public SwitchParameter RespectWorksheetHiddenRowsAndColumns { get; set; }

    /// <summary>Include worksheet row and column headings.</summary>
    [Parameter]
    public SwitchParameter IncludeSheetHeadings { get; set; }

    /// <summary>Number of leading rows treated as headers.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int? HeaderRowCount { get; set; }

    /// <summary>Maximum worksheet rows to read and render.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxRowsPerSheet { get; set; }

    /// <summary>Use bounded worksheet reads for large workbooks.</summary>
    [Parameter]
    public SwitchParameter UseBoundedWorksheetRead { get; set; }

    /// <summary>Text used when a worksheet cell is empty.</summary>
    [Parameter]
    public string? EmptyCellText { get; set; }

    /// <summary>Allow embedding fonts discovered on the current system.</summary>
    [Parameter]
    public SwitchParameter AllowSystemFontEmbedding { get; set; }

    /// <summary>Allow embedding installed fonts selected by a trusted workbook. System-font embedding must also be enabled.</summary>
    [Parameter]
    public SwitchParameter AllowDocumentFontEmbedding { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ExcelToPdfOptions();
        if (PdfOptions != null) options.PdfOptions = PdfOptions;
        if (!string.IsNullOrWhiteSpace(FontFamily)) options.FontFamily = FontFamily;
        if (PageSize.HasValue) options.PageSize = PageSize.Value;
        if (HasMargins()) {
            PageMargins defaults = PageMargins.Normal;
            options.Margins = new PageMargins(
                MarginLeft ?? defaults.Left,
                MarginTop ?? defaults.Top,
                MarginRight ?? defaults.Right,
                MarginBottom ?? defaults.Bottom);
        }
        if (WorksheetLayout.HasValue) options.WorksheetLayout = WorksheetLayout.Value;
        if (SheetName is { Length: > 0 }) options.SheetNames = new List<string>(SheetName);
        SetBoundSwitch(nameof(RespectWorkbookSheetVisibility), RespectWorkbookSheetVisibility, value => options.RespectWorkbookSheetVisibility = value);
        SetBoundSwitch(nameof(UseWorksheetPrintAreas), UseWorksheetPrintAreas, value => options.UseWorksheetPrintAreas = value);
        SetBoundSwitch(nameof(UseWorksheetPageSetup), UseWorksheetPageSetup, value => options.UseWorksheetPageSetup = value);
        SetBoundSwitch(nameof(UseWorksheetPrintTitleRows), UseWorksheetPrintTitleRows, value => options.UseWorksheetPrintTitleRows = value);
        SetBoundSwitch(nameof(UseWorksheetPageBreaks), UseWorksheetPageBreaks, value => options.UseWorksheetPageBreaks = value);
        SetBoundSwitch(nameof(UseWorksheetHeadersAndFooters), UseWorksheetHeadersAndFooters, value => options.UseWorksheetHeadersAndFooters = value);
        SetBoundSwitch(nameof(UseWorksheetHeaderFooterImages), UseWorksheetHeaderFooterImages, value => options.UseWorksheetHeaderFooterImages = value);
        SetBoundSwitch(nameof(UseWorksheetCellStyles), UseWorksheetCellStyles, value => options.UseWorksheetCellStyles = value);
        SetBoundSwitch(nameof(UseWorksheetHyperlinks), UseWorksheetHyperlinks, value => options.UseWorksheetHyperlinks = value);
        SetBoundSwitch(nameof(UseWorksheetImages), UseWorksheetImages, value => options.UseWorksheetImages = value);
        SetBoundSwitch(nameof(UseWorksheetCharts), UseWorksheetCharts, value => options.UseWorksheetCharts = value);
        if (ChartStyle != null) options.ChartStyle = ChartStyle;
        if (ChartLayout != null) options.ChartLayout = ChartLayout;
        SetBoundSwitch(nameof(UseWorksheetMergedCells), UseWorksheetMergedCells, value => options.UseWorksheetMergedCells = value);
        SetBoundSwitch(nameof(UseWorksheetColumnWidths), UseWorksheetColumnWidths, value => options.UseWorksheetColumnWidths = value);
        SetBoundSwitch(nameof(UseWorksheetRowHeights), UseWorksheetRowHeights, value => options.UseWorksheetRowHeights = value);
        SetBoundSwitch(nameof(RespectWorksheetHiddenRowsAndColumns), RespectWorksheetHiddenRowsAndColumns, value => options.RespectWorksheetHiddenRowsAndColumns = value);
        SetBoundSwitch(nameof(IncludeSheetHeadings), IncludeSheetHeadings, value => options.IncludeSheetHeadings = value);
        if (HeaderRowCount.HasValue) options.HeaderRowCount = HeaderRowCount.Value;
        if (MaxRowsPerSheet.HasValue) options.MaxRowsPerSheet = MaxRowsPerSheet.Value;
        SetBoundSwitch(nameof(UseBoundedWorksheetRead), UseBoundedWorksheetRead, value => options.UseBoundedWorksheetRead = value);
        if (EmptyCellText != null) options.EmptyCellText = EmptyCellText;
        SetBoundSwitch(nameof(AllowSystemFontEmbedding), AllowSystemFontEmbedding, value => options.ResourcePolicy.AllowSystemFontEmbedding = value);
        SetBoundSwitch(nameof(AllowDocumentFontEmbedding), AllowDocumentFontEmbedding, value => options.ResourcePolicy.AllowDocumentFontEmbedding = value);
        WriteObject(options);
    }

    private bool HasMargins() => MarginLeft.HasValue || MarginTop.HasValue || MarginRight.HasValue || MarginBottom.HasValue;
    private void SetBoundSwitch(string name, SwitchParameter value, System.Action<bool> setter) {
        if (MyInvocation.BoundParameters.ContainsKey(name)) setter(value.IsPresent);
    }
}
