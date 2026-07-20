using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

internal static class OneNoteVisualPdfRenderer {
    internal static PdfCore.PdfDocumentConversionResult Render(
        string sourceName,
        IReadOnlyList<OneNotePageReference> pages,
        OneNoteVisualPdfOptions? options,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OneNoteVisualPdfOptions effective = options?.Clone() ?? new OneNoteVisualPdfOptions();
        effective.Validate();
        string title = string.IsNullOrWhiteSpace(effective.Title) ? sourceName : effective.Title!;
        PdfCore.PdfDocument document = PdfCore.PdfDocument.Create().Meta(title, effective.Author, effective.Subject, effective.Keywords);
        var report = new PdfCore.PdfConversionReport();

        foreach (OneNotePageReference reference in pages) {
            cancellationToken.ThrowIfCancellationRequested();
            OneNotePageVisualSnapshot snapshot = OneNotePageRenderer.CreateSnapshot(reference.Page, effective.PageRendering);
            cancellationToken.ThrowIfCancellationRequested();
            OfficeRasterScaleLimit limit = OfficeRasterScaleLimiter.Resolve(
                snapshot.Drawing.Width, snapshot.Drawing.Height, effective.RasterScale, effective.MaximumRasterPixels);
            AddScaleDiagnostic(limit, effective.RasterScale, reference, report);
            var diagnostics = new List<OfficeImageExportDiagnostic>(snapshot.Diagnostics);
            var fallbackCodec = new OfficeRasterImageFallbackCodec(effective.PageRendering.ImageCodec, diagnostics, PageSource(reference));
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(snapshot.Drawing, new OfficeDrawingRasterRenderOptions {
                Scale = limit.Scale,
                Background = effective.PageRendering.BackgroundColor,
                ImageCodec = fallbackCodec,
                TextShapingProvider = effective.PageRendering.TextShapingProvider,
                TextShapingLanguage = effective.PageRendering.TextShapingLanguage,
                DiagnosticSink = diagnostics,
                DiagnosticSource = PageSource(reference),
                CancellationToken = cancellationToken
            });
            cancellationToken.ThrowIfCancellationRequested();
            byte[] png = OfficeRasterImageEncoder.Encode(raster, OfficeImageExportFormat.Png, effective.PageRendering.RasterEncoding);
            cancellationToken.ThrowIfCancellationRequested();
            string alternativeText = string.IsNullOrWhiteSpace(reference.Page.Title) ? "Untitled OneNote page" : reference.Page.Title;
            document.Page(page => page
                .Size(snapshot.Drawing.Width, snapshot.Drawing.Height)
                .Margin(0D)
                .Canvas(canvas => canvas.Image(png, 0D, 0D, snapshot.Drawing.Width, snapshot.Drawing.Height, alternativeText: alternativeText)));
            AddDiagnostics(reference, diagnostics, report);
        }

        cancellationToken.ThrowIfCancellationRequested();
        return new PdfCore.PdfDocumentConversionResult(document, report);
    }

    private static void AddScaleDiagnostic(
        OfficeRasterScaleLimit limit,
        double requestedScale,
        OneNotePageReference reference,
        PdfCore.PdfConversionReport report) {
        if (!limit.WasLimited) return;
        report.Add(new PdfCore.PdfConversionWarning(
            "OfficeIMO.OneNote.Pdf",
            "ONENOTE_PDF_RASTER_SCALE_LIMITED",
            PageSource(reference),
            "The raster scale was reduced from " + requestedScale.ToString("0.########", System.Globalization.CultureInfo.InvariantCulture) +
            " to " + limit.Scale.ToString("0.########", System.Globalization.CultureInfo.InvariantCulture) +
            " to respect MaximumRasterPixels."));
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
