using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private static void RenderRasterImages(OfficeRasterCanvas canvas, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics, System.Threading.CancellationToken cancellationToken) {
            foreach (ExcelVisualImage image in snapshot.Images) {
                cancellationToken.ThrowIfCancellationRequested();
                RenderRasterImage(canvas, image, options, diagnostics, cancellationToken);
            }
        }

        private static void AppendSvgImages(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            int index = 0;
            foreach (ExcelVisualImage image in snapshot.Images) {
                AppendSvgImage(builder, snapshot, image, options, diagnostics, ref index);
            }
        }

        private static void RenderRasterImage(OfficeRasterCanvas canvas, ExcelVisualImage image, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics, System.Threading.CancellationToken cancellationToken) {
            var drawingImage = new OfficeDrawingImage(image.Bytes, image.ContentType, CreateImageProjection(image, options.Scale));
            OfficeDrawingRasterRenderer.RenderImage(
                canvas,
                drawingImage,
                1D,
                new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, image.Source),
                options.MaximumRasterPixels,
                cancellationToken,
                diagnosticSource: image.Source);
        }

        private static void AppendSvgImage(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelVisualImage image, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics, ref int index) {
            double scale = options.Scale;
            var fallbackCodec = new OfficeRasterImageFallbackCodec(options.ImageCodec, diagnostics, image.Source);
            if (!OfficeSvgImageRenderer.TryCreateDataUri(image.ContentType, image.Bytes, image.Name, fallbackCodec, out string dataUri)) {
                return;
            }

            string clipId = "xl-image-clip-" + (++index).ToString(System.Globalization.CultureInfo.InvariantCulture);
            OfficeImageProjection projection = CreateImageProjection(image, scale);
            string safeUri = string.Empty;
            bool linked = image.HyperlinkUri != null
                && OfficeDrawingLinkPolicy.TryNormalize(image.HyperlinkUri.OriginalString, out safeUri);
            if (linked) {
                builder.Append("<a href=\"").Append(EscapeXml(safeUri)).Append("\">");
            } else if (image.HyperlinkUri != null) {
                diagnostics?.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ImageHyperlinkUnsupported,
                    "The picture hyperlink target is not supported by the safe SVG link policy; the image is retained without an interactive link.",
                    image.Source));
            }
            OfficeSvgImageRenderer.AppendImageInViewport(
                builder,
                dataUri,
                projection,
                clipId,
                new OfficeImagePlacement(0D, 0D, snapshot.Width * scale, snapshot.Height * scale));
            if (linked) builder.Append("</a>");
        }

        private static OfficeImageProjection CreateImageProjection(ExcelVisualImage image, double scale) =>
            OfficeImageRenderPlan.CreateTopLeft(
                image.SourceWidth > 0D ? image.SourceWidth : image.Width,
                image.SourceHeight > 0D ? image.SourceHeight : image.Height,
                image.X,
                image.Y,
                image.Width,
                image.Height,
                OfficeImageFit.Stretch,
                image.SourceCrop).ToVisibleProjection(
                image.RotationDegrees,
                flipHorizontal: image.FlipHorizontal,
                flipVertical: image.FlipVertical).Scale(scale);

    }
}
