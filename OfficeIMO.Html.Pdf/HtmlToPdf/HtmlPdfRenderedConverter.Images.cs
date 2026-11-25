using OfficeIMO.Drawing;
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
            alternativeText: string.IsNullOrWhiteSpace(visual.AlternativeText) ? null : visual.AlternativeText,
            logicalOrder: visual.SemanticSourceOrder);
        if (fragmentLink) {
            canvas.LinkToNamedDestination(
                MapNamedDestination(visual.LinkUri!.Substring(1)),
                global::OfficeIMO.OfficeLength.Points(visual.X * PointsPerCssPixel),
                global::OfficeIMO.OfficeLength.Points(visual.Y * PointsPerCssPixel),
                global::OfficeIMO.OfficeLength.Points(visual.Width * PointsPerCssPixel),
                global::OfficeIMO.OfficeLength.Points(visual.Height * PointsPerCssPixel),
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
        canvas.SourceClip(area.X, area.Y, area.Width, area.Height, clipped => {
            foreach (OfficeImagePlacement tile in pattern.GetTilePlacements(visual.MaximumTileCount)) {
                cancellationToken.ThrowIfCancellationRequested();
                clipped.ImageShared(imageResource, tile.X, tile.Y, tile.Width, tile.Height);
            }
        });
    }

    // Cache lifetime follows one export: a codec failure or a resource decision must not
    // leak into another export that uses different providers or limits.
    private const double DefaultEffectRasterDpi = 300D;

    internal sealed class PdfImageResourceCache {
        private readonly Dictionary<byte[], Dictionary<string, PreparedImage>> _resources = new();
        private readonly HashSet<PreparedImage> _reportedLosses = new();
        private readonly OfficeRasterDecodeOptions _decodeOptions;
        private PdfCore.PdfConversionReport? _conversionReport;

        private sealed class PreparedImage {
            internal PreparedImage(PdfCore.PdfCanvasImageResource? resource, string? frameLoss) {
                Resource = resource;
                FrameLoss = frameLoss;
            }

            internal PdfCore.PdfCanvasImageResource? Resource { get; }
            internal string? FrameLoss { get; }
        }

        internal IOfficeRasterImageCodec? ImageCodec => _decodeOptions.ImageCodec;
        internal long MaximumPixels => _decodeOptions.MaximumDecodedPixels;

        /// <summary>Raster scale, per CSS pixel, for paint that native PDF cannot carry.</summary>
        /// <remarks>A requested <see cref="OfficeImageExportOptions.TargetDpi"/> sets the density; otherwise print density.</remarks>
        internal double EffectRasterScale { get; }

        internal PdfImageResourceCache(HtmlPdfExportOptions options) {
            EffectRasterScale = (options.TargetDpi ?? DefaultEffectRasterDpi) / HtmlRenderOptions.CssPixelsPerInch;
            _decodeOptions = new OfficeRasterDecodeOptions {
                ImageCodec = options.ImageCodec,
                MaximumDecodedPixels = System.Math.Min(options.MaximumRasterPixels, 50_000_000L)
            };
        }

        internal void BeginSection(PdfCore.PdfConversionReport conversionReport, CancellationToken cancellationToken) {
            _conversionReport = conversionReport;
            _decodeOptions.CancellationToken = cancellationToken;
            _reportedLosses.Clear();
        }

        internal PdfCore.PdfCanvasImageResource? GetOrCreate(byte[] encodedBytes, string contentType) {
            _decodeOptions.CancellationToken.ThrowIfCancellationRequested();
            if (!_resources.TryGetValue(encodedBytes, out var byContentType)) {
                byContentType = new Dictionary<string, PreparedImage>(System.StringComparer.OrdinalIgnoreCase);
                _resources.Add(encodedBytes, byContentType);
            }
            if (!byContentType.TryGetValue(contentType, out PreparedImage? image)) {
                bool success = TryPrepare(encodedBytes, contentType, out byte[] prepared, out string? frameLoss);
                image = new PreparedImage(success ? PdfCore.PdfCanvasImageResource.Create(prepared) : null, frameLoss);
                byContentType.Add(contentType, image);
            }
            if (image.FrameLoss != null && _reportedLosses.Add(image)) {
                _conversionReport!.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.Html.Pdf", HtmlPdfDiagnosticCodes.ImageStaticFrameSelected, "html-image-resource",
                    image.FrameLoss, PdfCore.PdfConversionWarningSeverity.Warning, OfficeConversionLossKind.Omission));
            }
            return image.Resource;
        }

        private bool TryPrepare(byte[] bytes, string contentType, out byte[] pdfBytes, out string? frameLoss) {
            frameLoss = null;
            OfficeImageFormat format = OfficeImageInfo.FromMimeType(contentType);
            string extension = OfficeImageInfo.GetDefaultExtension(format);
            if (OfficeImageReader.TryIdentify(bytes, extension, _decodeOptions.CancellationToken, out OfficeImageInfo identified))
                format = identified.Format;
            if (format == OfficeImageFormat.Png || format == OfficeImageFormat.Jpeg) {
                if (!OfficeRasterContainerInspector.TryInspectForDecode(
                        bytes, _decodeOptions, out var container, out _, out _, out _) ||
                    container == null || container.Format != format) {
                    pdfBytes = System.Array.Empty<byte>();
                    return false;
                }
                if (format == OfficeImageFormat.Png && (container.IsAnimated || container.Count > 1)) {
                    frameLoss = "The default static PNG image was retained; APNG animation and additional frames were discarded.";
                }
                pdfBytes = bytes;
                return true;
            }
            bool converted = OfficeImagePngConverter.TryConvertToPng(bytes, _decodeOptions, out pdfBytes, out var info);
            if (converted && (info.AnimationDiscarded || info.FramesOrPagesDiscarded)) {
                frameLoss = info.Diagnostic ?? "The first image frame was selected; additional frames or animation were discarded.";
            }
            return converted;
        }
    }
}
