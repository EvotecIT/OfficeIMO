using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

internal static class OneNoteVisualPdfRenderer {
    internal const string ImageRasterizedCode = "ONENOTE_PDF_IMAGE_RASTERIZED";

    internal static PdfCore.PdfDocumentConversionResult Render(
        string sourceName,
        IReadOnlyList<OneNotePageReference> pages,
        OneNoteVisualPdfExportOptions? options,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OneNoteVisualPdfExportOptions effective = options?.Clone() ?? new OneNoteVisualPdfExportOptions();
        effective.Validate();
        effective.PdfOptions.UseContentStreamCompressionByDefault();
        var report = new PdfCore.PdfConversionReport();
        var drawings = new List<OfficeDrawing>(pages.Count);
        var fonts = new OfficeFontFaceCollection();
        foreach (OneNotePageReference reference in pages) {
            cancellationToken.ThrowIfCancellationRequested();
            string source = PageSource(reference);
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            OneNotePageRenderingOptions pageRendering = effective.PageRendering.Clone();
            pageRendering.ImageSubstitution = (byte[] bytes, string? contentType, out byte[] substitute, out string substituteContentType) =>
                TryEmbedUnsupportedImageAsPng(bytes, contentType, pageRendering, diagnostics, source, cancellationToken, out substitute, out substituteContentType);
            OneNotePageVisualSnapshot snapshot = OneNotePageRenderer.CreateSnapshot(reference.Page, pageRendering);
            diagnostics.InsertRange(0, snapshot.Diagnostics);
            drawings.Add(snapshot.Drawing);
            fonts.AddRange(snapshot.Drawing.Fonts);
            AddDiagnostics(reference, diagnostics, report);
        }

        cancellationToken.ThrowIfCancellationRequested();
        effective.PdfOptions.UseRenderingProfile(new OfficeRenderingProfile(
            "onenote-visual-pages",
            fonts: fonts,
            textShapingProvider: effective.PageRendering.TextShapingProvider,
            textShapingLanguage: effective.PageRendering.TextShapingLanguage), OfficeRenderingProfileApplyMode.Overlay);
        effective.PdfOptions.UseDrawingTextFonts(drawings, cancellationToken);
        effective.PdfOptions.ReportDiagnosticsTo(report, "OfficeIMO.OneNote.Pdf");
        string title = string.IsNullOrWhiteSpace(effective.Title) ? sourceName : effective.Title!;
        PdfCore.PdfDocument document = PdfCore.PdfDocument.CreateForImportedLayout(effective.PdfOptions).Meta(title, effective.Author, effective.Subject, effective.Keywords);
        var origin = global::OfficeIMO.OfficeLength.Points(0D);
        foreach (OfficeDrawing drawing in drawings) {
            cancellationToken.ThrowIfCancellationRequested();
            var width = global::OfficeIMO.OfficeLength.Points(drawing.Width);
            var height = global::OfficeIMO.OfficeLength.Points(drawing.Height);
            document.Compose(pdf => pdf.Page(page => page
                .Size(width, height)
                .Margin(origin)
                .Canvas(canvas => canvas.Drawing(drawing, origin, origin, width, height))));
        }

        cancellationToken.ThrowIfCancellationRequested();
        return new PdfCore.PdfDocumentConversionResult(document, report);
    }

    // The PDF writer embeds JPEG and PNG directly and normalizes formats its own decoder reads.
    // Anything else is decoded through the caller's codec (or replaced by the shared visible
    // placeholder) and embedded as PNG for this picture only; the rest of the page stays vector.
    private static bool TryEmbedUnsupportedImageAsPng(
        byte[] bytes,
        string? contentType,
        OneNotePageRenderingOptions rendering,
        ICollection<OfficeImageExportDiagnostic> diagnostics,
        string source,
        CancellationToken cancellationToken,
        out byte[] substitute,
        out string substituteContentType) {
        cancellationToken.ThrowIfCancellationRequested();
        substitute = bytes;
        substituteContentType = contentType ?? string.Empty;
        if (PdfCore.PdfDocument.TryPrepareImageBytes(bytes, out _, out _, out _, out string? unsupportedReason)) return false;

        var codec = new OfficeRasterImageFallbackCodec(rendering.ImageCodec, diagnostics, source);
        codec.TryDecode(bytes, contentType, out OfficeRasterImage? raster);
        substitute = OfficeRasterImageEncoder.Encode(raster!, OfficeImageExportFormat.Png, rendering.RasterEncoding);
        substituteContentType = "image/png";
        string format = string.IsNullOrWhiteSpace(contentType) ? "An embedded picture" : "An embedded " + contentType + " picture";
        diagnostics.Add(new OfficeImageExportDiagnostic(
            OfficeImageExportDiagnosticSeverity.Info,
            ImageRasterizedCode,
            format + " has no native PDF image form and was embedded as PNG." +
            (string.IsNullOrWhiteSpace(unsupportedReason) ? string.Empty : " " + unsupportedReason),
            source));
        return true;
    }

    private static void AddDiagnostics(
        OneNotePageReference reference,
        IReadOnlyList<OfficeImageExportDiagnostic> diagnostics,
        PdfCore.PdfConversionReport report) {
        foreach (OfficeImageExportDiagnostic diagnostic in diagnostics) {
            report.Add(new PdfCore.PdfConversionWarning(
                "OfficeIMO.OneNote.Pdf",
                diagnostic.Code,
                diagnostic.Source ?? PageSource(reference),
                diagnostic.Message,
                diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error
                    ? PdfCore.PdfConversionWarningSeverity.Error
                    : diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Warning
                        ? PdfCore.PdfConversionWarningSeverity.Warning
                        : PdfCore.PdfConversionWarningSeverity.Information));
        }
    }

    private static string PageSource(OneNotePageReference reference) =>
        reference.SectionPath + "/page[" + reference.Index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
}
