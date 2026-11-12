using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Pdf;

internal static class OneNoteVisualPdfRenderer {
    internal const string ImageRasterizedCode = "ONENOTE_PDF_IMAGE_RASTERIZED";
    internal const string GlyphSubstitutedCode = "ONENOTE_PDF_GLYPH_SUBSTITUTED";

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
        List<PageSnapshot> snapshots = CreateSnapshots(pages, effective, textSubstitution: null, cancellationToken);
        var fonts = new OfficeFontFaceCollection();
        foreach (PageSnapshot snapshot in snapshots) fonts.AddRange(snapshot.Drawing.Fonts);

        cancellationToken.ThrowIfCancellationRequested();
        effective.PdfOptions.UseRenderingProfile(new OfficeRenderingProfile(
            "onenote-visual-pages",
            fonts: fonts,
            textShapingProvider: effective.PageRendering.TextShapingProvider,
            textShapingLanguage: effective.PageRendering.TextShapingLanguage), OfficeRenderingProfileApplyMode.Overlay);
        effective.PdfOptions.UseDrawingTextFonts(snapshots.Select(snapshot => snapshot.Drawing), cancellationToken);

        // A character no installed font covers would fail the PDF text preflight. Lay the pages out again
        // with those characters replaced by a covered compatibility form or a visible missing-glyph mark.
        var substitutions = new SortedDictionary<string, string>(StringComparer.Ordinal);
        if (effective.PdfOptions.HasUncoveredDrawingText(snapshots.Select(snapshot => snapshot.Drawing), cancellationToken)) {
            PdfCore.PdfCreateOptions pdfOptions = effective.PdfOptions;
            snapshots = CreateSnapshots(pages, effective, (text, font) => pdfOptions.SubstituteUncoveredDrawingText(
                text, font.FamilyName, font.IsBold, font.IsItalic, (original, replacement) => substitutions[original] = replacement), cancellationToken);
        }
        foreach (PageSnapshot snapshot in snapshots) AddDiagnostics(snapshot.Reference, snapshot.Diagnostics, report);
        if (substitutions.Count > 0) {
            report.Add(new PdfCore.PdfConversionWarning(
                "OfficeIMO.OneNote.Pdf",
                GlyphSubstitutedCode,
                sourceName,
                "No available font covers " + string.Join(", ", substitutions.Select(pair => Describe(pair.Key) + " (drawn as " + Describe(pair.Value) + ")")) +
                ". Install or configure a font with these glyphs to keep them.",
                PdfCore.PdfConversionWarningSeverity.Warning));
        }
        List<OfficeDrawing> drawings = snapshots.Select(snapshot => snapshot.Drawing).ToList();
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

    private sealed class PageSnapshot {
        internal PageSnapshot(OneNotePageReference reference, OfficeDrawing drawing, IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
            Reference = reference;
            Drawing = drawing;
            Diagnostics = diagnostics;
        }

        internal OneNotePageReference Reference { get; }
        internal OfficeDrawing Drawing { get; }
        internal IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }
    }

    private static List<PageSnapshot> CreateSnapshots(
        IReadOnlyList<OneNotePageReference> pages,
        OneNoteVisualPdfExportOptions effective,
        Func<string, OfficeFontInfo, string>? textSubstitution,
        CancellationToken cancellationToken) {
        var snapshots = new List<PageSnapshot>(pages.Count);
        foreach (OneNotePageReference reference in pages) {
            cancellationToken.ThrowIfCancellationRequested();
            string source = PageSource(reference);
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            OneNotePageRenderingOptions pageRendering = effective.PageRendering.Clone();
            pageRendering.ImageSubstitution = (byte[] bytes, string? contentType, out byte[] substitute, out string substituteContentType) =>
                TryEmbedUnsupportedImageAsPng(bytes, contentType, pageRendering, diagnostics, source, cancellationToken, out substitute, out substituteContentType);
            pageRendering.TextSubstitution = textSubstitution;
            OneNotePageVisualSnapshot snapshot = OneNotePageRenderer.CreateSnapshot(reference.Page, pageRendering);
            diagnostics.InsertRange(0, snapshot.Diagnostics);
            snapshots.Add(new PageSnapshot(reference, snapshot.Drawing, diagnostics));
        }
        return snapshots;
    }

    private static string Describe(string text) {
        var builder = new System.Text.StringBuilder();
        for (int index = 0; index < text.Length; index++) {
            int scalar = char.IsHighSurrogate(text[index]) && index + 1 < text.Length ? char.ConvertToUtf32(text[index], text[++index]) : text[index];
            if (builder.Length > 0) builder.Append(' ');
            builder.Append("U+").Append(scalar.ToString("X4", System.Globalization.CultureInfo.InvariantCulture));
        }
        return builder.ToString();
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
