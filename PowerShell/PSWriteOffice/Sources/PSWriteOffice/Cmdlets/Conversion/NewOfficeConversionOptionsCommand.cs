using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.PowerPoint;
using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word.Pdf;
using DrawPdfOptions = OfficeIMO.OpenDocument.Odg.Pdf.OdgToPdfOptions;
using VisioPdfOptions = OfficeIMO.Visio.Pdf.VisioToPdfOptions;
using XpsPdfOptions = OfficeIMO.Xps.XpsToPdfOptions;
using DjVuPdfOptions = OfficeIMO.DjVu.Pdf.DjVuToPdfOptions;
using OfficeIMO.Workflows;

namespace PSWriteOffice.Cmdlets.Conversion;

/// <summary>Creates discoverable renderer and reconstruction settings for general conversion jobs.</summary>
/// <para>Format-specific options apply only to their matching routes.</para>
/// <example>
///   <summary>Keep workbook charts in a strict conversion job.</summary>
///   <code>$options = New-OfficeConversionOptions -RequireNoLoss -ExcelOptions (New-OfficeExcelPdfOptions -UseWorksheetCharts)
/// ConvertTo-OfficeDocument -InputPath ./report.xlsx -OutputPath ./report.pdf -ConversionOptions $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeConversionOptions")]
public sealed class NewOfficeConversionOptionsCommand : PSCmdlet {
    /// <summary>Word-to-PDF settings from New-OfficeWordPdfOptions.</summary>
    [Parameter]
    public WordPdfExportOptions? WordOptions { get; set; }

    /// <summary>Excel-to-PDF settings from New-OfficeExcelPdfOptions.</summary>
    [Parameter]
    public ExcelPdfExportOptions? ExcelOptions { get; set; }

    /// <summary>PowerPoint-to-PDF settings from New-OfficePowerPointPdfOptions.</summary>
    [Parameter]
    public PowerPointPdfExportOptions? PowerPointOptions { get; set; }

    /// <summary>HTML-to-PDF settings from New-OfficeHtmlPdfOptions.</summary>
    [Parameter]
    public HtmlToPdfOptions? HtmlOptions { get; set; }

    /// <summary>Markdown-to-PDF settings from New-OfficeMarkdownPdfOptions.</summary>
    [Parameter]
    public MarkdownPdfExportOptions? MarkdownOptions { get; set; }

    /// <summary>RTF-to-PDF settings from New-OfficeRtfPdfOptions.</summary>
    [Parameter]
    public RtfToPdfOptions? RtfOptions { get; set; }

    /// <summary>Literal-text settings from New-OfficeTextPdfOptions.</summary>
    [Parameter]
    public PdfTextImportOptions? TextOptions { get; set; }

    /// <summary>ODG/FODG drawing projection and PDF options; accepts a property hashtable.</summary>
    [Parameter]
    public DrawPdfOptions? DrawOptions { get; set; }

    /// <summary>Visio diagram-page projection settings; accepts a property hashtable.</summary>
    [Parameter]
    public VisioPdfOptions? VisioOptions { get; set; }

    /// <summary>Native XPS preservation and PDF output settings.</summary>
    [Parameter]
    public XpsPdfOptions? XpsOptions { get; set; }

    /// <summary>CHM topic selection and conversion budgets.</summary>
    [Parameter]
    public OfficeIMO.Chm.ChmConversionOptions? ChmOptions { get; set; }

    /// <summary>DjVu scanned-page PDF and stored-text settings.</summary>
    [Parameter]
    public DjVuPdfOptions? DjVuOptions { get; set; }

    /// <summary>Native Publisher decoding and projection limits.</summary>
    [Parameter]
    public OfficeIMO.Publisher.PublisherReadOptions? PublisherReadOptions { get; set; }

    /// <summary>Shared PDF settings for recovered Publisher pages.</summary>
    [Parameter]
    public PdfCreateOptions? PublisherPdfOptions { get; set; }

    /// <summary>Losslessly compress generated PDF output.</summary>
    [Parameter]
    public SwitchParameter CompressPdfOutput { get; set; }

    /// <summary>Workbook worksheet-canvas or flowing-table layout.</summary>
    [Parameter]
    public ExcelPdfWorksheetLayoutMode? WorksheetLayout { get; set; }

    /// <summary>Semantic or positioned-review HTML reconstructed from PDF.</summary>
    [Parameter]
    public PdfHtmlProfile? PdfHtmlProfile { get; set; }

    /// <summary>Reject any reduction reported by the selected route.</summary>
    [Parameter]
    public SwitchParameter RequireNoLoss { get; set; }

    /// <summary>Explicitly accept known legacy DOC import reductions.</summary>
    [Parameter]
    public SwitchParameter AllowLegacyImportLoss { get; set; }

    /// <summary>Page ranges for PDF source reconstruction.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>PDF-to-Word reconstruction or rendered-page strategy.</summary>
    [Parameter]
    public WordPdfImportMode? PdfWordMode { get; set; }

    /// <summary>PDF-to-PowerPoint reconstruction or rendered-page strategy.</summary>
    [Parameter]
    public PowerPointPdfImportMode? PdfPowerPointMode { get; set; }

    /// <summary>Resolution of rendered PDF pages, when applicable.</summary>
    [Parameter]
    [ValidateRange(36d, 600d)]
    public double? RasterDpi { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new OfficeConversionOptions {
            Word = WordOptions, Excel = ExcelOptions?.Clone(), PowerPoint = PowerPointOptions,
            Html = HtmlOptions, Markdown = MarkdownOptions, PlainText = TextOptions,
            RegisteredSettings = new OfficeWorkflowRouteSettings {
                Rtf = RtfOptions, Draw = DrawOptions, Visio = VisioOptions, Xps = XpsOptions,
                Chm = ChmOptions, DjVu = DjVuOptions, PublisherRead = PublisherReadOptions, PublisherPdf = PublisherPdfOptions
            },
            CompressPdfOutput = CompressPdfOutput.IsPresent, FailOnLoss = RequireNoLoss.IsPresent,
            WordLoad = new OfficeIMO.Word.WordLoadOptions { FailOnLoss = !AllowLegacyImportLoss.IsPresent },
            PdfLoad = PageRange == null ? null : new PdfLoadOptions { PageSelection = PdfPageSelection.Parse(PageRange) },
            PdfWord = PdfWordMode == null && RasterDpi == null ? null : new WordPdfImportOptions { Mode = PdfWordMode ?? WordPdfImportMode.EditableContent, Dpi = RasterDpi ?? 144D },
            PdfPowerPoint = PdfPowerPointMode == null && RasterDpi == null ? null : new PowerPointPdfImportOptions { Mode = PdfPowerPointMode ?? PowerPointPdfImportMode.Auto, Dpi = RasterDpi ?? 144D },
            PdfHtml = PdfHtmlProfile == null ? null : new PdfToHtmlOptions { Profile = PdfHtmlProfile.Value }
        };
        if (WorksheetLayout is { } layout) { options.Excel ??= new ExcelPdfExportOptions(); options.Excel.WorksheetLayout = layout; }
        WriteObject(options);
    }
}
