using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO;
using PSWriteOffice.Foundation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates discoverable Excel-to-PDF conversion options for Convert-OfficeDocument -To Pdf.</summary>
/// <example>
/// <summary>Convert chosen sheets to PDF</summary>
/// <code>$pdf = New-OfficeExcelPdfOptions -Sheet Summary, Orders -UseWorksheetPrintAreas
/// Convert-OfficeDocument -Path .\Reports\Q3-Sales.xlsx -OutputPath .\Reports\Q3-Sales.pdf -Options $pdf</code>
/// <para>Converts only the Summary and Orders sheets, honouring the print areas set in the workbook.</para>
/// </example>
/// <example>
/// <summary>Convert a folder of reports with page setup</summary>
/// <code>$pdf = New-OfficeExcelPdfOptions -PageSize A4 -MarginLeft 1.5cm -MarginRight 1.5cm -UseWorksheetPageSetup -UseWorksheetHeadersAndFooters -UseWorksheetPrintTitleRows -RespectWorkbookSheetVisibility
/// Get-ChildItem -Path .\Reports -Filter *.xlsx |
///     Convert-OfficeDocument -To Pdf -OutputDirectory .\Pdf -Options $pdf</code>
/// <para>Converts every workbook in the folder to A4 PDFs that use each sheet's page setup, headers and repeated titles.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcelPdfOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(ExcelPdfExportOptions))]
public sealed class NewOfficeExcelPdfOptionsCommand : PSCmdlet {
    /// <summary>Underlying low-level OfficeIMO PDF options.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeIMO.Pdf.PdfCreateOptions? PdfOptions { get; set; }

    /// <summary>Default font family used when the workbook does not specify one.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Font { get; set; }

    /// <summary>PDF page size.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficePageSizeTransformation]
    [ArgumentCompleter(typeof(OfficePdfPageSizeArgumentCompleter))]
    public PdfPageSize? PageSize { get; set; }

    /// <summary>Left page margin in PDF points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MarginLeft { get; set; }

    /// <summary>Top page margin in PDF points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MarginTop { get; set; }

    /// <summary>Right page margin in PDF points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MarginRight { get; set; }

    /// <summary>Bottom page margin in PDF points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? MarginBottom { get; set; }

    /// <summary>Controls how worksheet content is laid out on PDF pages.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelPdfWorksheetLayoutMode? WorksheetLayout { get; set; }

    /// <summary>Worksheet names, wildcard patterns or one-based positions. Live worksheet inputs are captured by name for reusable saved conversions.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector[]? Sheet { get; set; }

    /// <summary>Exclude workbook sheets marked hidden.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter RespectWorkbookSheetVisibility { get; set; }

    /// <summary>Honor worksheet print areas.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetPrintAreas { get; set; }

    /// <summary>Honor worksheet page setup.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetPageSetup { get; set; }

    /// <summary>Honor worksheet rows configured to repeat on printed pages.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetPrintTitleRows { get; set; }

    /// <summary>Honor worksheet page breaks.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetPageBreaks { get; set; }

    /// <summary>Render worksheet headers and footers.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetHeadersAndFooters { get; set; }

    /// <summary>Render images referenced by worksheet headers and footers.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetHeaderFooterImages { get; set; }

    /// <summary>Render worksheet cell styles.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetCellStyles { get; set; }

    /// <summary>Render worksheet hyperlinks.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetHyperlinks { get; set; }

    /// <summary>Render worksheet images.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetImages { get; set; }

    /// <summary>Render worksheet charts.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetCharts { get; set; }

    /// <summary>Chart visual style override.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeChartStyle? ChartStyle { get; set; }

    /// <summary>Chart layout override.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeChartLayout? ChartLayout { get; set; }

    /// <summary>Render merged worksheet cells.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetMergedCells { get; set; }

    /// <summary>Honor worksheet column widths.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetColumnWidths { get; set; }

    /// <summary>Honor worksheet row heights.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseWorksheetRowHeights { get; set; }

    /// <summary>Exclude hidden worksheet rows and columns.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter RespectWorksheetHiddenRowsAndColumns { get; set; }

    /// <summary>Include worksheet row and column headings.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeSheetHeadings { get; set; }

    /// <summary>Number of leading rows treated as headers.</summary>
    [Parameter(ParameterSetName = "Default")]
    [ValidateRange(0, int.MaxValue)]
    public int? HeaderRowCount { get; set; }

    /// <summary>Maximum worksheet rows to read and render.</summary>
    [Parameter(ParameterSetName = "Default")]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxRowsPerSheet { get; set; }

    /// <summary>Use bounded worksheet reads for large workbooks.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter UseBoundedWorksheetRead { get; set; }

    /// <summary>Text used when a worksheet cell is empty.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? EmptyCellText { get; set; }

    /// <summary>Allow embedding fonts discovered on the current system.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowSystemFontEmbedding { get; set; }

    /// <summary>Allow embedding installed fonts selected by a trusted workbook. System-font embedding must also be enabled.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter AllowDocumentFontEmbedding { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ExcelPdfExportOptions();
        if (PdfOptions != null) options.PdfOptions = PdfOptions;
        if (!string.IsNullOrWhiteSpace(Font)) options.FontFamily = Font;
        if (PageSize.HasValue) options.PageSize = PageSize.Value;
        if (HasMargins()) {
            PdfPageMargins defaults = PdfPageMargins.Normal;
            options.Margins = new PdfPageMargins(
                MarginLeft.HasValue ? MarginLeft.Value : defaults.Left,
                MarginTop.HasValue ? MarginTop.Value : defaults.Top,
                MarginRight.HasValue ? MarginRight.Value : defaults.Right,
                MarginBottom.HasValue ? MarginBottom.Value : defaults.Bottom);
        }
        if (WorksheetLayout.HasValue) options.WorksheetLayout = WorksheetLayout.Value;

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
        options.Sheets = Sheet?.Select(selector => new OfficeSheetSelector(selector.Name, selector.Position)).ToArray();
        WriteObject(options);
    }

    private bool HasMargins() => MarginLeft.HasValue || MarginTop.HasValue || MarginRight.HasValue || MarginBottom.HasValue;
    private void SetBoundSwitch(string name, SwitchParameter value, System.Action<bool> setter) {
        if (MyInvocation.BoundParameters.ContainsKey(name)) setter(value.IsPresent);
    }
}
