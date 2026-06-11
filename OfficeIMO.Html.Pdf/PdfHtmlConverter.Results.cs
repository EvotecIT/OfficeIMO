using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverter {
    /// <summary>
    /// Converts PDF bytes to HTML and returns a machine-readable export summary.
    /// </summary>
    public static PdfHtmlConversionResult ToHtmlResult(byte[] pdf, PdfHtmlSaveOptions? options = null) {
        if (pdf == null) {
            throw new ArgumentNullException(nameof(pdf));
        }

        options ??= new PdfHtmlSaveOptions();
        options.ResetExportState();
        return RenderLogicalDocumentResult(LoadFullLogical(pdf, options), options, applyPageRanges: true);
    }

    /// <summary>
    /// Converts a PDF file to HTML and returns a machine-readable export summary.
    /// </summary>
    public static PdfHtmlConversionResult ToHtmlResult(string path, PdfHtmlSaveOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) {
            throw new ArgumentException("PDF path cannot be empty.", nameof(path));
        }

        options ??= new PdfHtmlSaveOptions();
        options.ResetExportState();
        return RenderLogicalDocumentResult(LoadFullLogical(path, options), options, applyPageRanges: true);
    }

    /// <summary>
    /// Converts PDF stream content to HTML and returns a machine-readable export summary.
    /// </summary>
    public static PdfHtmlConversionResult ToHtmlResult(Stream stream, PdfHtmlSaveOptions? options = null) {
        if (stream == null) {
            throw new ArgumentNullException(nameof(stream));
        }

        options ??= new PdfHtmlSaveOptions();
        options.ResetExportState();
        return RenderLogicalDocumentResult(LoadFullLogical(stream, options), options, applyPageRanges: true);
    }

    /// <summary>
    /// Renders an already parsed PDF document as HTML and returns a machine-readable export summary.
    /// </summary>
    public static PdfHtmlConversionResult ToHtmlResult(this PdfCore.PdfReadDocument document, PdfHtmlSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new PdfHtmlSaveOptions();
        options.ResetExportState();
        return RenderLogicalDocumentResult(LoadFullLogical(document, options), options, applyPageRanges: true);
    }

    /// <summary>
    /// Renders an already loaded logical PDF model as HTML and returns a machine-readable export summary.
    /// </summary>
    public static PdfHtmlConversionResult ToHtmlResult(this PdfCore.PdfLogicalDocument document, PdfHtmlSaveOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new PdfHtmlSaveOptions();
        options.ResetExportState();
        return RenderLogicalDocumentResult(document, options, applyPageRanges: true);
    }

    private static PdfHtmlConversionResult RenderLogicalDocumentResult(PdfCore.PdfLogicalDocument document, PdfHtmlSaveOptions options, bool applyPageRanges) {
        IReadOnlyList<PdfCore.PdfLogicalPage> pages = applyPageRanges
            ? GetRenderPages(document, options)
            : document.Pages;
        string html = options.Profile switch {
            PdfHtmlProfile.Semantic => RenderSemanticDocument(document, pages, options),
            PdfHtmlProfile.PositionedReview => RenderPositionedReviewDocument(document, pages, options),
            _ => throw new ArgumentOutOfRangeException(nameof(options.Profile), options.Profile, "Unsupported PDF HTML profile.")
        };
        return new PdfHtmlConversionResult(html, BuildExportSummary(document, pages, options), options.ConversionReport);
    }

    private static PdfCore.PdfLogicalDocument LoadFullLogical(byte[] pdf, PdfHtmlSaveOptions options) {
        return PdfCore.PdfLogicalDocument.Load(pdf, options.LayoutOptions);
    }

    private static PdfCore.PdfLogicalDocument LoadFullLogical(string path, PdfHtmlSaveOptions options) {
        return PdfCore.PdfLogicalDocument.Load(path, options.LayoutOptions);
    }

    private static PdfCore.PdfLogicalDocument LoadFullLogical(Stream stream, PdfHtmlSaveOptions options) {
        return PdfCore.PdfLogicalDocument.Load(stream, options.LayoutOptions);
    }

    private static PdfCore.PdfLogicalDocument LoadFullLogical(PdfCore.PdfReadDocument document, PdfHtmlSaveOptions options) {
        return PdfCore.PdfLogicalDocument.From(document, options.LayoutOptions);
    }

    private static PdfHtmlExportSummary BuildExportSummary(PdfCore.PdfLogicalDocument document, IReadOnlyList<PdfCore.PdfLogicalPage> pages, PdfHtmlSaveOptions options) {
        int textBlockCount = 0;
        int headingCount = 0;
        int listItemCount = 0;
        int tableCount = 0;
        int imageCount = 0;
        int imagePlacementCount = 0;
        int linkCount = 0;
        int formWidgetCount = 0;
        var pageNumbers = new int[pages.Count];
        var formFields = new HashSet<PdfCore.PdfFormField>();

        for (int i = 0; i < pages.Count; i++) {
            PdfCore.PdfLogicalPage page = pages[i];
            pageNumbers[i] = page.PageNumber;
            textBlockCount += page.TextBlocks.Count;
            headingCount += page.Headings.Count;
            listItemCount += page.ListItems.Count;
            tableCount += page.Tables.Count;
            imageCount += page.Images.Count;
            linkCount += page.Links.Count;
            formWidgetCount += page.FormWidgets.Count;
            for (int widgetIndex = 0; widgetIndex < page.FormWidgets.Count; widgetIndex++) {
                formFields.Add(page.FormWidgets[widgetIndex].Field);
            }

            for (int imageIndex = 0; imageIndex < page.Images.Count; imageIndex++) {
                imagePlacementCount += page.Images[imageIndex].PlacementCount;
            }
        }

        PdfHtmlProfileContract contract = PdfHtmlProfileContracts.Get(options.Profile);
        return new PdfHtmlExportSummary(
            options.Profile,
            contract.Id,
            pageNumbers,
            document.PageCount,
            pages.Count,
            textBlockCount,
            headingCount,
            listItemCount,
            tableCount,
            imageCount,
            imagePlacementCount,
            linkCount,
            formFields.Count,
            formWidgetCount,
            options.ConversionReport.Warnings.Count,
            options.EmitDocumentShell,
            options.ImageExportMode,
            contract.FidelityContract,
            contract.UnsupportedScope);
    }
}
