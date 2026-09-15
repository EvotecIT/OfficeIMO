using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

public static partial class PdfHtmlConverterExtensions {
    private static readonly System.Threading.AsyncLocal<Action?> ImagePayloadHashObserver =
        new System.Threading.AsyncLocal<Action?>();

    internal static Action? ImagePayloadHashObserverForTesting {
        get => ImagePayloadHashObserver.Value;
        set => ImagePayloadHashObserver.Value = value;
    }

    private const string PageAppearancePrefix = "<img class=\"pdf-page-appearance\" aria-hidden=\"true\" alt=\"\" draggable=\"false\" decoding=\"sync\" src=\"data:image/svg+xml;base64,";
    private const string PageAppearanceSuffix = "\" style=\"position:absolute;inset:0;width:100%;height:100%;user-select:none;pointer-events:none\" />\n";

    private static bool TryAppendPageAppearance(StringBuilder builder, PdfCore.PdfLogicalPage page,
        int renderIndex, PdfToHtmlOptions options) {
        if (options.VisualSource is null ||
            page.FormWidgets.Count > 0 ||
            page.Analysis.RestrictLogicalProjectionToReadingOrder ||
            page.TextBlocks.Any(block => block.Spans.Count == 0)) return false;

        var token = options.CancellationToken;
        token.ThrowIfCancellationRequested();
        if (page.Images.Any(image => image.Placements.Any(placement => {
                PdfCore.PdfImagePlacementImportAssessment assessment =
                    PdfCore.PdfImagePlacementImportPolicy.Analyze(page, image, placement);
                return !assessment.CanImport && !assessment.IsSuppressed;
            }))) {
            return ReportUnsafeImageAppearanceFallback(options);
        }
        bool hasImportableImages = page.Images.Any(image => image.Placements.Any(placement =>
            PdfCore.PdfImagePlacementImportPolicy.Analyze(page, image, placement).CanImport));
        if (hasImportableImages &&
            (!options.IncludeImagePlaceholders ||
             options.ImageExportMode != PdfHtmlImageExportMode.EmbeddedDataUri)) return false;
        long pixelBudget = 100_000;
        foreach (var image in page.Images) {
            if (!image.Placements.Any(placement =>
                    PdfCore.PdfImagePlacementImportPolicy.Analyze(page, image, placement).CanImport)) continue;
            PdfCore.PdfExtractedImage sourceImage = image.SourceImage;
            if (options.MaxEmbeddedImageBytes.HasValue &&
                sourceImage.Bytes.LongLength > options.MaxEmbeddedImageBytes.Value) return false;
            if (!TryConsumePageAppearanceImageBudget(sourceImage, ref pixelBudget, token)) {
                return ReportImageAppearanceFallback(options);
            }
        }
        var source = options.VisualSource.GetReadDocument(options.VisualSource.ReadOptions, token);
        var sourcePage = source.Pages[page.PageNumber - 1];
        // The shared PDF projection owns clipping, paint order, embedded fonts,
        // images and paths. Never rebuild those semantics from logical blocks.
        var drawing = sourcePage.ToDrawing(token);
        if (ContainsUnaccountedDrawingImagePayload(drawing, page, token)) {
            return ReportUnaccountedImageAppearanceFallback(options);
        }
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

    private static bool TryConsumePageAppearanceImageBudget(
        PdfCore.PdfExtractedImage image,
        ref long remainingPixels,
        System.Threading.CancellationToken cancellationToken) {
        if (!image.Interpolate) {
            long declaredPixels = (long)image.Width * image.Height;
            if (declaredPixels <= 0L || declaredPixels > remainingPixels) return false;
        }
        if (!CanRenderPageAppearanceImage(image, remainingPixels, cancellationToken, out long decodedPixels)) {
            return false;
        }
        remainingPixels -= decodedPixels;
        return true;
    }

    internal static bool CanRenderPageAppearanceImagesWithinBudgetForTesting(
        IReadOnlyList<PdfCore.PdfExtractedImage> images,
        long maximumDecodedPixels,
        System.Threading.CancellationToken cancellationToken = default) {
        long remainingPixels = maximumDecodedPixels;
        for (int index = 0; index < images.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryConsumePageAppearanceImageBudget(images[index], ref remainingPixels, cancellationToken)) {
                return false;
            }
        }
        return true;
    }

    private static bool ReportImageAppearanceFallback(PdfToHtmlOptions options) {
        AddWarning(options, "PageAppearanceImageFallback",
            "The page uses positioned images and text to avoid excessive SVG image expansion. Compare its appearance with the source PDF.",
            PdfCore.PdfConversionWarningSeverity.Warning);
        return false;
    }

    private static bool ReportUnsafeImageAppearanceFallback(PdfToHtmlOptions options) {
        AddWarning(options, "PageAppearanceUnsafeImageFallback",
            "The page appearance SVG was not emitted because it could expose image pixels hidden by PDF clipping, transparency, or paint effects. Positioned HTML fallback was used instead.",
            PdfCore.PdfConversionWarningSeverity.Information);
        return false;
    }

    private static bool ReportUnaccountedImageAppearanceFallback(PdfToHtmlOptions options) {
        AddWarning(options, "PageAppearanceUnaccountedImageOmitted",
            "The page appearance SVG was not emitted because it contains image payloads that cannot be correlated with safely importable logical images. Positioned HTML fallback omitted those images.",
            PdfCore.PdfConversionWarningSeverity.Warning,
            OfficeConversionLossKind.Omission);
        return false;
    }

    private static bool ContainsUnaccountedDrawingImagePayload(
        OfficeDrawing drawing,
        PdfCore.PdfLogicalPage page,
        System.Threading.CancellationToken cancellationToken) {
        var safePayloadCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var payloadKeys = new Dictionary<byte[], string>();
        foreach (PdfCore.PdfLogicalImage image in page.Images) {
            cancellationToken.ThrowIfCancellationRequested();
            string key = GetImagePayloadKey(image.SourceImage.EncodedBytes, payloadKeys, cancellationToken);
            foreach (PdfCore.PdfImagePlacement placement in image.Placements) {
                if (!PdfCore.PdfImagePlacementImportPolicy.Analyze(page, image, placement).CanImport) continue;
                safePayloadCounts[key] = safePayloadCounts.TryGetValue(key, out int count) ? count + 1 : 1;
            }
        }

        return ContainsUnaccountedDrawingImagePayload(drawing, safePayloadCounts, payloadKeys, cancellationToken);
    }

    private static bool ContainsUnaccountedDrawingImagePayload(
        OfficeDrawing drawing,
        IDictionary<string, int> safePayloadCounts,
        IDictionary<byte[], string> payloadKeys,
        System.Threading.CancellationToken cancellationToken) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            cancellationToken.ThrowIfCancellationRequested();
            switch (element) {
                case OfficeDrawingImage image:
                    if (!ConsumeSafeImagePayload(image.EncodedBytes, safePayloadCounts, payloadKeys, cancellationToken)) return true;
                    break;
                case OfficeDrawingImagePattern imagePattern:
                    if (!ConsumeSafeImagePayload(imagePattern.EncodedBytes, safePayloadCounts, payloadKeys, cancellationToken)) return true;
                    break;
                case OfficeDrawingGroup group:
                    if (ContainsUnaccountedDrawingImagePayload(group.InnerDrawing, safePayloadCounts, payloadKeys, cancellationToken)) return true;
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    if (ContainsUnaccountedDrawingImagePayload(effectGroup.InnerDrawing, safePayloadCounts, payloadKeys, cancellationToken) ||
                        effectGroup.SoftMask != null &&
                        ContainsUnaccountedDrawingImagePayload(effectGroup.SoftMask.InnerDrawing, safePayloadCounts, payloadKeys, cancellationToken)) return true;
                    break;
                case OfficeDrawingTilingPattern tilingPattern:
                    if (ContainsUnaccountedDrawingImagePayload(tilingPattern.InnerTile, safePayloadCounts, payloadKeys, cancellationToken)) return true;
                    break;
            }
        }

        return false;
    }

    private static bool ConsumeSafeImagePayload(
        byte[] bytes,
        IDictionary<string, int> safePayloadCounts,
        IDictionary<byte[], string> payloadKeys,
        System.Threading.CancellationToken cancellationToken) {
        string key = GetImagePayloadKey(bytes, payloadKeys, cancellationToken);
        if (!safePayloadCounts.TryGetValue(key, out int count) || count <= 0) return false;
        safePayloadCounts[key] = count - 1;
        return true;
    }

    private static string GetImagePayloadKey(
        byte[] bytes,
        IDictionary<byte[], string> payloadKeys,
        System.Threading.CancellationToken cancellationToken) {
        if (payloadKeys.TryGetValue(bytes, out string? cached)) return cached;
        cancellationToken.ThrowIfCancellationRequested();
        ImagePayloadHashObserver.Value?.Invoke();
        cancellationToken.ThrowIfCancellationRequested();
        using SHA256 sha256 = SHA256.Create();
        string key = bytes.Length.ToString(CultureInfo.InvariantCulture) + ":" +
            Convert.ToBase64String(sha256.ComputeHash(bytes));
        cancellationToken.ThrowIfCancellationRequested();
        payloadKeys.Add(bytes, key);
        return key;
    }
}
