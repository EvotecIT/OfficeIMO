using OfficeIMO.Drawing;
using System.Collections.Generic;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static bool AddImage(PdfCore.PdfPageCanvas canvas, HtmlRenderImage visual, PdfImageResourceCache imageResources, bool suppressLink) {
        PdfCore.PdfCanvasImageResource? imageResource = imageResources.GetOrCreate(
            visual.EncodedBytes, visual.ContentType);
        if (imageResource == null) return false;
        PdfCore.PdfImageStyle? style = visual.SourceCrop.HasCrop
            ? new PdfCore.PdfImageStyle {
                SourceCrop = new PdfCore.PdfImageSourceCrop(
                    visual.SourceCrop.Left,
                    visual.SourceCrop.Top,
                    visual.SourceCrop.Right,
                    visual.SourceCrop.Bottom)
            }
            : null;
        bool fragmentLink = !suppressLink && IsFragmentLink(visual.LinkUri);
        canvas.ImageShared(
            imageResource,
            visual.X * PointsPerCssPixel,
            visual.Y * PointsPerCssPixel,
            visual.Width * PointsPerCssPixel,
            visual.Height * PointsPerCssPixel,
            style,
            linkUri: suppressLink || fragmentLink ? null : visual.LinkUri,
            linkContents: suppressLink || visual.LinkUri == null || fragmentLink ? null : visual.Source,
            alternativeText: string.IsNullOrWhiteSpace(visual.AlternativeText) ? null : visual.AlternativeText);
        if (fragmentLink) {
            canvas.LinkToNamedDestination(
                MapNamedDestination(visual.LinkUri!.Substring(1)),
                visual.X * PointsPerCssPixel,
                visual.Y * PointsPerCssPixel,
                visual.Width * PointsPerCssPixel,
                visual.Height * PointsPerCssPixel,
                visual.Source);
        }
        return true;
    }

    private static void AddImagePattern(PdfCore.PdfPageCanvas canvas, HtmlRenderImagePattern visual, PdfImageResourceCache imageResources, CancellationToken cancellationToken) {
        PdfCore.PdfCanvasImageResource? imageResource = imageResources.GetOrCreate(
            visual.EncodedBytes, visual.ContentType);
        if (imageResource == null) return;
        OfficeImagePatternLayout pattern = visual.Pattern.Scale(PointsPerCssPixel);
        OfficeImagePlacement area = pattern.Area;
        canvas.Clip(area.X, area.Y, area.Width, area.Height, clipped => {
            foreach (OfficeImagePlacement tile in pattern.GetTilePlacements(visual.MaximumTileCount)) {
                cancellationToken.ThrowIfCancellationRequested();
                clipped.ImageShared(imageResource, tile.X, tile.Y, tile.Width, tile.Height);
            }
        });
    }

    // Cache lifetime follows one export: a codec failure or a resource decision must not
    // leak into another export that uses different providers or limits.
    private sealed class PdfImageResourceCache {
        private readonly Dictionary<byte[], Dictionary<string, PdfCore.PdfCanvasImageResource?>> _resources = new();
        private readonly OfficeRasterDecodeOptions _decodeOptions;
        private readonly PdfCore.PdfConversionReport _conversionReport;

        internal IOfficeRasterImageCodec? ImageCodec => _decodeOptions.ImageCodec;
        internal long MaximumPixels => _decodeOptions.MaximumDecodedPixels;

        internal PdfImageResourceCache(HtmlToPdfOptions options, PdfCore.PdfConversionReport conversionReport, CancellationToken cancellationToken) {
            _conversionReport = conversionReport;
            _decodeOptions = new OfficeRasterDecodeOptions {
                ImageCodec = options.ImageCodec,
                MaximumDecodedPixels = System.Math.Min(options.MaximumRasterPixels, 50_000_000L),
                CancellationToken = cancellationToken
            };
        }

        internal PdfCore.PdfCanvasImageResource? GetOrCreate(byte[] encodedBytes, string contentType) {
            _decodeOptions.CancellationToken.ThrowIfCancellationRequested();
            if (!_resources.TryGetValue(encodedBytes, out var byContentType)) {
                byContentType = new Dictionary<string, PdfCore.PdfCanvasImageResource?>(System.StringComparer.OrdinalIgnoreCase);
                _resources.Add(encodedBytes, byContentType);
            }
            if (byContentType.TryGetValue(contentType, out var cached)) return cached;

            PdfCore.PdfCanvasImageResource? resource = null;
            if (TryPrepare(encodedBytes, contentType, out byte[] prepared))
                resource = PdfCore.PdfCanvasImageResource.Create(prepared);
            byContentType.Add(contentType, resource);
            return resource;
        }

        private bool TryPrepare(byte[] bytes, string contentType, out byte[] pdfBytes) {
            OfficeImageFormat format = OfficeImageInfo.FromMimeType(contentType);
            string extension = OfficeImageInfo.GetDefaultExtension(format);
            if (OfficeImageReader.TryIdentify(bytes, extension, _decodeOptions.CancellationToken, out OfficeImageInfo identified))
                format = identified.Format;
            if (format == OfficeImageFormat.Png || format == OfficeImageFormat.Jpeg) {
                if (format == OfficeImageFormat.Png && OfficeRasterContainerInspector.TryInspectForDecode(
                    bytes, _decodeOptions, out var container, out _, out _, out _) && container != null &&
                    (container.IsAnimated || container.Count > 1)) {
                    _conversionReport.Add(new PdfCore.PdfConversionWarning(
                        "OfficeIMO.Html.Pdf", HtmlPdfDiagnosticCodes.ImageStaticFrameSelected, "html-image-resource",
                        "The default static PNG image was retained; APNG animation and additional frames were discarded.",
                        PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission));
                }
                pdfBytes = bytes;
                return true;
            }
            bool converted = OfficeImagePngConverter.TryConvertToPng(bytes, _decodeOptions, out pdfBytes, out var info);
            if (converted && (info.AnimationDiscarded || info.FramesOrPagesDiscarded)) {
                _conversionReport.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.Html.Pdf", HtmlPdfDiagnosticCodes.ImageStaticFrameSelected, "html-image-resource",
                    info.Diagnostic ?? "The first image frame was selected; other frames or animation were not retained in the static PDF.",
                    PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission));
            }
            return converted;
        }
    }
}
