using System;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
    private static bool TryAppendPageAppearance(StringBuilder builder, PdfCore.PdfLogicalPage page,
        int renderIndex, PdfToHtmlOptions options) {
        if (options.VisualSource is null || !options.IncludeImagePlaceholders ||
            options.ImageExportMode != PdfHtmlImageExportMode.EmbeddedDataUri ||
            page.FormWidgets.Count > 0 ||
            page.Analysis.RestrictLogicalProjectionToReadingOrder ||
            page.TextBlocks.Any(block => block.Spans.Count == 0) ||
            (options.MaxEmbeddedImageBytes.HasValue && page.Images.Any(image =>
                image.SourceImage.Bytes.LongLength > options.MaxEmbeddedImageBytes.Value))) return false;

        var token = options.CancellationToken;
        token.ThrowIfCancellationRequested();
        long pixelBudget = 100_000;
        foreach (var image in page.Images) {
            if (image.SourceImage.Interpolate) continue;
            pixelBudget -= (long)image.SourceImage.Width * image.SourceImage.Height;
            if (pixelBudget < 0) return ReportImageAppearanceFallback(options);
        }
        var source = options.VisualSource.GetReadDocument(options.VisualSource.ReadOptions, token);
        var sourcePage = source.Pages[page.PageNumber - 1];
        // The shared PDF projection owns clipping, paint order, embedded fonts,
        // images and paths. Never rebuild those semantics from logical blocks.
        var drawing = sourcePage.ToDrawing(token);
        long remaining = options.MaximumOutputCharacters.HasValue
            ? options.MaximumOutputCharacters.Value - (long)builder.Length
            : int.MaxValue;
        if (remaining <= 0) throw new InvalidOperationException("Generated HTML exceeded its output limit.");
        byte[] svg;
        try {
            svg = OfficeDrawingSvgExporter.ToSvgBytes(drawing, 1D, OfficeSvgSizeUnit.Point,
                imageCodec: null, resourceIdPrefix: "pdf-page-" + renderIndex.ToString(CultureInfo.InvariantCulture) + "-",
                maximumUtf8Bytes: remaining * 3L, cancellationToken: token);
        } catch (OfficeSvgImageVectorizationLimitException) {
            token.ThrowIfCancellationRequested();
            return ReportImageAppearanceFallback(options);
        }
        token.ThrowIfCancellationRequested();
        builder.Append("<div class=\"pdf-page-appearance\" style=\"position:absolute;inset:0\">");
        builder.Append(Encoding.UTF8.GetString(svg));
        builder.AppendLine("</div>");
        foreach (var diagnostic in sourcePage.GetRenderCapabilityDiagnostics(token)) {
            AddWarning(options, diagnostic.Code, "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + ": " + diagnostic.Message,
                PdfCore.PdfConversionWarningSeverity.Warning);
        }
        if (drawing.Fonts.Faces.Count == 0 && page.TextBlocks.Count > 0) {
            AddWarning(options, "PageAppearanceBrowserFonts",
                "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + ": fonts are not embedded. Browser font substitution may change text appearance.",
                PdfCore.PdfConversionWarningSeverity.Warning);
        }
        return true;
    }

    private static bool ReportImageAppearanceFallback(PdfToHtmlOptions options) {
        AddWarning(options, "PageAppearanceImageFallback",
            "The page uses positioned images and text to avoid excessive SVG image expansion. Compare its appearance with the source PDF.",
            PdfCore.PdfConversionWarningSeverity.Warning);
        return false;
    }
}
