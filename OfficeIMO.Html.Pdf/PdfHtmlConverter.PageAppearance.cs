using System;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
    private const string PageAppearancePrefix = "<img class=\"pdf-page-appearance\" aria-hidden=\"true\" alt=\"\" draggable=\"false\" decoding=\"sync\" src=\"data:image/svg+xml;base64,";
    private const string PageAppearanceSuffix = "\" style=\"position:absolute;inset:0;width:100%;height:100%;user-select:none;pointer-events:none\" />\n";

    private static bool TryAppendPageAppearance(StringBuilder builder, PdfCore.PdfLogicalPage page,
        int renderIndex, PdfToHtmlOptions options) {
        if (options.VisualSource is null ||
            page.FormWidgets.Count > 0 ||
            page.Analysis.RestrictLogicalProjectionToReadingOrder ||
            page.TextBlocks.Any(block => block.Spans.Count == 0) ||
            (page.Images.Count > 0 &&
                (!options.IncludeImagePlaceholders ||
                 options.ImageExportMode != PdfHtmlImageExportMode.EmbeddedDataUri ||
                 (options.MaxEmbeddedImageBytes.HasValue && page.Images.Any(image =>
                     image.SourceImage.Bytes.LongLength > options.MaxEmbeddedImageBytes.Value))))) return false;

        var token = options.CancellationToken;
        token.ThrowIfCancellationRequested();
        long pixelBudget = 100_000;
        foreach (var image in page.Images) {
            PdfCore.PdfExtractedImage sourceImage = image.SourceImage;
            if (!sourceImage.Interpolate) {
                long declaredPixels = (long)sourceImage.Width * sourceImage.Height;
                if (declaredPixels <= 0L || declaredPixels > pixelBudget) {
                    return ReportImageAppearanceFallback(options);
                }
            }
            if (!CanRenderPageAppearanceImage(sourceImage, pixelBudget, token, out long decodedPixels)) {
                return ReportImageAppearanceFallback(options);
            }
            if (!sourceImage.Interpolate) pixelBudget -= decodedPixels;
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
        long base64Capacity = remaining - PageAppearancePrefix.Length - PageAppearanceSuffix.Length;
        long maximumSvgBytes = base64Capacity >= 4L ? base64Capacity / 4L * 3L : 0L;
        if (maximumSvgBytes <= 0L) throw new InvalidOperationException("Generated HTML exceeded its output limit.");
        byte[] svg;
        try {
            svg = OfficeDrawingSvgExporter.ToSvgBytes(drawing, 1D, OfficeSvgSizeUnit.Point,
                imageCodec: null, resourceIdPrefix: "pdf-page-" + renderIndex.ToString(CultureInfo.InvariantCulture) + "-",
                maximumUtf8Bytes: maximumSvgBytes, cancellationToken: token);
        } catch (OfficeSvgImageVectorizationLimitException) {
            token.ThrowIfCancellationRequested();
            return ReportImageAppearanceFallback(options);
        }
        token.ThrowIfCancellationRequested();
        // Keep the visual SVG in an image document. Inline SVG text participates in
        // browser find, selection, and copy even when aria-hidden, which would expose
        // the same content a second time beside the logical text overlay.
        builder.Append(PageAppearancePrefix);
        builder.Append(Convert.ToBase64String(svg));
        builder.Append(PageAppearanceSuffix);
        foreach (var diagnostic in sourcePage.GetRenderCapabilityDiagnostics(token)) {
            AddWarning(options, diagnostic.Code, "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + ": " + diagnostic.Message,
                diagnostic.SupportLevel == PdfCore.PdfRenderSupportLevel.Unsupported
                    ? PdfCore.PdfConversionWarningSeverity.Warning
                    : PdfCore.PdfConversionWarningSeverity.Information,
                diagnostic.SupportLevel == PdfCore.PdfRenderSupportLevel.Unsupported
                    ? OfficeConversionLossKind.Omission
                    : OfficeConversionLossKind.Approximation);
        }
        if (drawing.Fonts.Faces.Count == 0 && page.TextBlocks.Count > 0) {
            AddWarning(options, "PageAppearanceBrowserFonts",
                "Page " + page.PageNumber.ToString(CultureInfo.InvariantCulture) + ": fonts are not embedded. Browser font substitution may change text appearance.",
                PdfCore.PdfConversionWarningSeverity.Warning);
        }
        return true;
    }

    private static bool CanRenderPageAppearanceImage(
        PdfCore.PdfExtractedImage image,
        long maximumDecodedPixels,
        System.Threading.CancellationToken cancellationToken,
        out long decodedPixels) {
        decodedPixels = 0L;
        byte[] imageBytes = image.Bytes;
        if (image.Interpolate && OfficeSvgImageRenderer.TryResolveEmbeddableContentType(
                image.MimeType, imageBytes, image.FileExtension, out _)) {
            return true;
        }
        if (maximumDecodedPixels <= 0L || !OfficeRasterImageDecoder.TryDecode(
            imageBytes,
            new OfficeRasterDecodeOptions {
                MaximumDecodedPixels = maximumDecodedPixels,
                CancellationToken = cancellationToken
            },
            out OfficeRasterImage? raster,
            out _) || raster is null) return false;
        decodedPixels = (long)raster.Width * raster.Height;
        return decodedPixels > 0L && decodedPixels <= maximumDecodedPixels;
    }

    private static bool ReportImageAppearanceFallback(PdfToHtmlOptions options) {
        AddWarning(options, "PageAppearanceImageFallback",
            "The page uses positioned images and text to avoid excessive SVG image expansion. Compare its appearance with the source PDF.",
            PdfCore.PdfConversionWarningSeverity.Warning);
        return false;
    }
}
