using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.OpenDocument.Odg.Pdf;
using OfficeIMO.Visio.Pdf;
using OfficeIMO.Xps;
using OfficeIMO.DjVu.Pdf;
using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

/// <summary>Route-specific settings captured before a conversion starts. Unspecified values retain each route's defaults.</summary>
public sealed class OfficeWorkflowConversionOptions {
    /// <summary>Runtime password for encrypted DOCX, XLSX and PPTX inputs. Never serialized into checkpoint settings.</summary>
    [JsonIgnore]
    public string? SourcePassword { get; set; }
    /// <summary>Existing Word renderer settings for DOC and DOCX input.</summary>
    public WordToPdfOptions? Word { get; set; }
    /// <summary>Existing Excel renderer settings for XLSX input.</summary>
    public ExcelToPdfOptions? Excel { get; set; }
    /// <summary>Existing presentation renderer settings for PPTX input.</summary>
    public PowerPointToPdfOptions? PowerPoint { get; set; }
    /// <summary>Existing HTML renderer settings. Workflow resource access remains scoped to the source.</summary>
    public HtmlToPdfOptions? Html { get; set; }
    /// <summary>Existing Markdown renderer settings.</summary>
    public MarkdownToPdfOptions? Markdown { get; set; }
    /// <summary>Existing RTF renderer settings.</summary>
    public RtfToPdfOptions? Rtf { get; set; }
    /// <summary>Draw page projection and PDF settings for ODG/FODG input.</summary>
    public OdgToPdfOptions? Draw { get; set; }
    /// <summary>Cached diagram-page projection and PDF settings for VSDX, VDX and VTX input. Explicit settings require DiagramPages mode.</summary>
    public VisioToPdfOptions? Visio { get; set; }
    /// <summary>Rejects source or PDF-stage conversion losses for diagram and DjVu PDF routes before publication.</summary>
    public bool RequireNoLoss { get; set; }
    /// <summary>Native XPS/OpenXPS semantic preservation and PDF output settings.</summary>
    public XpsToPdfOptions? Xps { get; set; }
    /// <summary>Scanned-page PDF and stored-text settings. Workflow OCR uses separate explicit OCR operations.</summary>
    public DjVuToPdfOptions? DjVu { get; set; }
    /// <summary>Literal-text settings, valid only for the TXT-to-PDF route.</summary>
    public PdfPlainTextOptions? PlainText { get; set; }
    /// <summary>Known legacy DOC import loss blocks output unless explicitly accepted.</summary>
    public OfficeConversionLossPolicy LegacyDocLossPolicy { get; set; } = OfficeConversionLossPolicy.Block;
    /// <summary>Optional source PDF page ranges, for example 1-3,5. Other source formats do not accept this setting.</summary>
    public string? PageRanges { get; set; }
    /// <summary>PDF-to-Word reconstruction or rendered-page strategy.</summary>
    public PdfWordImportMode? WordMode { get; set; }
    /// <summary>PDF-to-PowerPoint reconstruction or rendered-page strategy.</summary>
    public PdfPowerPointImportMode? PowerPointMode { get; set; }
    /// <summary>Resolution for visual PDF-to-Word or PDF-to-PowerPoint pages, from 36 through 600 DPI.</summary>
    public double? RasterDpi { get; set; }
    /// <summary>Worksheet canvas or flowing table layout for Excel-to-PDF conversion.</summary>
    public ExcelPdfWorksheetLayoutMode? WorksheetLayout { get; set; }
    /// <summary>Semantic or positioned review HTML for PDF-to-HTML conversion.</summary>
    public PdfHtmlProfile? HtmlProfile { get; set; }
    /// <summary>Applies verified lossless compression after conversion to PDF. Other destination formats reject true.</summary>
    public bool CompressPdfOutput { get; set; }

    /// <summary>Creates an independent settings copy.</summary>
    public OfficeWorkflowConversionOptions Clone() {
        var copy = (OfficeWorkflowConversionOptions)MemberwiseClone();
        copy.PlainText = PlainText?.Clone();
        copy.Word = Word?.Clone();
        copy.Excel = Excel?.Clone();
        copy.PowerPoint = PowerPoint?.Clone();
        copy.Html = Html?.ClonePdf();
        copy.Markdown = Markdown?.Clone();
        copy.Rtf = Rtf?.Clone();
        copy.Draw = Draw?.Clone();
        if (Visio != null) copy.Visio = new VisioToPdfOptions {
            Mode = Visio.Mode, SourceName = Visio.SourceName,
            DrawingOptions = Visio.DrawingOptions?.Clone(), PdfOptions = Visio.PdfOptions?.Clone(),
            // Semantic settings are rejected by workflow validation; preserve them for that check.
            VisioOptions = Visio.VisioOptions, ProjectionOptions = Visio.ProjectionOptions
        };
        copy.Xps = Xps?.Clone();
        copy.DjVu = DjVu?.Clone();
        return copy;
    }

    /// <summary>Selects settings applicable to one route in a mixed-format batch.</summary>
    public OfficeWorkflowConversionOptions ForRoute(string routeId) {
        OfficeWorkflowRoute route = OfficeWorkflowCatalog.FindExecutable(routeId)
            ?? throw new ArgumentException("Choose an executable conversion route.", nameof(routeId));
        return ForRoute(route);
    }

    /// <summary>Selects settings using the executable route captured from the configured runner.</summary>
    internal OfficeWorkflowConversionOptions ForRoute(OfficeWorkflowRoute route) {
        string routeId = route.Id;
        var copy = Clone();
        if (routeId is not "docx-pdf" and not "xlsx-pdf" and not "pptx-pdf") copy.SourcePassword = null;
        if (routeId is not "doc-pdf" and not "docx-pdf") copy.Word = null;
        if (routeId != "xlsx-pdf") { copy.Excel = null; copy.WorksheetLayout = null; }
        if (routeId != "pptx-pdf") copy.PowerPoint = null;
        if (routeId != "html-pdf") copy.Html = null;
        if (routeId != "markdown-pdf") copy.Markdown = null;
        if (routeId != "rtf-pdf") copy.Rtf = null;
        if (routeId != "odg-pdf") copy.Draw = null;
        if (routeId != "visio-pdf") copy.Visio = null;
        if (routeId is not "odg-pdf" and not "visio-pdf" and not "djvu-pdf") copy.RequireNoLoss = false;
        if (routeId != "djvu-pdf") copy.DjVu = null;
        if (routeId != "xps-pdf") copy.Xps = null;
        if (routeId != "txt-pdf") copy.PlainText = null;
        if (routeId != "doc-pdf") copy.LegacyDocLossPolicy = OfficeConversionLossPolicy.Block;
        if (!route.SupportsPageSelection) copy.PageRanges = null;
        if (routeId != "pdf-docx") copy.WordMode = null;
        if (routeId != "pdf-pptx") copy.PowerPointMode = null;
        if (copy.WordMode != PdfWordImportMode.VisualPages && copy.PowerPointMode is not
            PdfPowerPointImportMode.VisualPages and not PdfPowerPointImportMode.HybridVisualAndEditableTables) copy.RasterDpi = null;
        if (routeId != "pdf-html") copy.HtmlProfile = null;
        if (!route.SupportsPdfCompression) copy.CompressPdfOutput = false;
        return copy;
    }

    internal OfficeWorkflowConversionOptions Snapshot(OfficeWorkflowRoute route) {
        OfficeWorkflowConversionOptions copy = Clone();
        if (copy.SourcePassword != null && route.Id is not "docx-pdf" and not "xlsx-pdf" and not "pptx-pdf")
            throw new ArgumentException("A source password is supported for DOCX, XLSX and PPTX conversion.");
        if ((copy.Word != null && route.Id is not "doc-pdf" and not "docx-pdf") ||
            (copy.Excel != null && route.Id != "xlsx-pdf") || (copy.PowerPoint != null && route.Id != "pptx-pdf") ||
            (copy.Html != null && route.Id != "html-pdf") || (copy.Markdown != null && route.Id != "markdown-pdf") ||
            (copy.Rtf != null && route.Id != "rtf-pdf") || (copy.Draw != null && route.Id != "odg-pdf") ||
            (copy.Visio != null && route.Id != "visio-pdf") || (copy.Xps != null && route.Id != "xps-pdf") || (copy.DjVu != null && route.Id != "djvu-pdf"))
            throw new ArgumentException("Renderer settings must match the selected conversion route.");
        if (copy.DjVu?.OcrEngine != null)
            throw new ArgumentException("Use the native asynchronous DjVu converter or the explicit OCR workflow for OCR engines.");
        if (copy.Visio != null && (copy.Visio.Mode != VisioPdfProjectionMode.DiagramPages ||
            copy.Visio.VisioOptions != null || copy.Visio.ProjectionOptions != null))
            throw new ArgumentException("Visio workflow conversion requires DiagramPages settings.");
        if (copy.Html?.ResourceResolver != null)
            throw new ArgumentException("Workflow HTML resources use the scoped source resolver. Use the native HTML adapter for runtime custom resolvers.");
        if (copy.RequireNoLoss && route.Id is not "odg-pdf" and not "visio-pdf" and not "djvu-pdf")
            throw new ArgumentException("Lossless acceptance requires a diagram or DjVu PDF route.");
        if (copy.PlainText != null && route.Id != "txt-pdf") throw new ArgumentException("Plain-text settings require TXT-to-PDF conversion.");
        if (copy.LegacyDocLossPolicy is not OfficeConversionLossPolicy.Block and not OfficeConversionLossPolicy.Allow ||
            (route.Id != "doc-pdf" && copy.LegacyDocLossPolicy != OfficeConversionLossPolicy.Block))
            throw new ArgumentException("Legacy import loss can be accepted only for DOC-to-PDF conversion.");
        if (!string.IsNullOrWhiteSpace(copy.PageRanges)) {
            if (!route.SupportsPageSelection) throw new ArgumentException("Page ranges require a PDF input route.");
            if (copy.PageRanges.Length > 2048) throw new ArgumentException("The page selection is too long.");
            _ = PdfPageSelection.Parse(copy.PageRanges);
        }
        if (copy.WordMode.HasValue && (route.Id != "pdf-docx" || !Enum.IsDefined(copy.WordMode.Value)))
            throw new ArgumentException("Choose a supported PDF-to-Word import mode.");
        if (copy.PowerPointMode.HasValue && (route.Id != "pdf-pptx" || !Enum.IsDefined(copy.PowerPointMode.Value)))
            throw new ArgumentException("Choose a supported PDF-to-PowerPoint import mode.");
        if (copy.WorksheetLayout.HasValue && (route.Id != "xlsx-pdf" || !Enum.IsDefined(copy.WorksheetLayout.Value)))
            throw new ArgumentException("Worksheet layout is valid only for Excel-to-PDF conversion.");
        if (copy.HtmlProfile.HasValue && (route.Id != "pdf-html" || !Enum.IsDefined(copy.HtmlProfile.Value)))
            throw new ArgumentException("HTML layout is valid only for PDF-to-HTML conversion.");
        if (copy.CompressPdfOutput && !route.SupportsPdfCompression)
            throw new ArgumentException("PDF compression requires a PDF output route.");
        bool visual = copy.WordMode == PdfWordImportMode.VisualPages || copy.PowerPointMode is
            PdfPowerPointImportMode.VisualPages or PdfPowerPointImportMode.HybridVisualAndEditableTables;
        if (copy.RasterDpi is double dpi && (!visual || !double.IsFinite(dpi) || dpi < 36 || dpi > 600))
            throw new ArgumentException("Rendering resolution requires visual PDF pages and must be between 36 and 600 DPI.");
        return copy;
    }

    internal PdfOptions? GetOutputPdfOptions() => Word?.PdfOptions ?? Excel?.PdfOptions ?? PowerPoint?.PdfOptions
        ?? Html?.PdfOptions ?? Markdown?.PdfOptions ?? Rtf?.PdfOptions ?? Draw?.PdfOptions ?? Visio?.PdfOptions ?? PlainText?.PdfOptions ?? Xps?.PdfOptions ?? DjVu?.PdfOptions;

    internal PdfReadOptions? CreateReadOptions() => string.IsNullOrWhiteSpace(PageRanges) ? null
        : new PdfReadOptions { PageSelection = PdfPageSelection.Parse(PageRanges) };
}
