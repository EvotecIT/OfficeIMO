using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private static void RenderRasterImages(OfficeRasterCanvas canvas, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            foreach (ExcelVisualImage image in snapshot.Images) {
                RenderRasterImage(canvas, image, options, diagnostics);
            }
        }

        private static void AppendSvgImages(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            int index = 0;
            foreach (ExcelVisualImage image in snapshot.Images) {
                AppendSvgImage(builder, snapshot, image, options, diagnostics, ref index);
            }
        }

        private static void RenderRasterImage(OfficeRasterCanvas canvas, ExcelVisualImage image, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            double scale = options.Scale;
            if (image.DetectedFormat != OfficeImageFormat.Png) {
                diagnostics?.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ImageRasterFormatUnsupported,
                    "Worksheet image format '" + image.DetectedFormat + "' cannot be rasterized to PNG by the dependency-free image renderer yet.",
                    image.Source));
                return;
            }

            if (!OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster) || raster == null) {
                diagnostics?.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ImagePngDecodeUnavailable,
                    "Worksheet PNG image bytes could not be decoded for PNG output.",
                    image.Source));
                return;
            }

            canvas.DrawImage(raster, CreateImageProjection(image, scale));
        }

        private static void AppendSvgImage(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelVisualImage image, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics, ref int index) {
            double scale = options.Scale;
            if (!OfficeSvgImageRenderer.TryGetEmbeddableContentType(image.DetectedFormat, out string contentType)) {
                diagnostics?.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ImageSvgFormatUnsupported,
                    "Worksheet image format '" + image.DetectedFormat + "' cannot be embedded reliably in SVG output by the dependency-free image renderer yet.",
                    image.Source));
                return;
            }

            string clipId = "xl-image-clip-" + (++index).ToString(System.Globalization.CultureInfo.InvariantCulture);
            OfficeImageProjection projection = CreateImageProjection(image, scale);
            OfficeSvgImageRenderer.AppendImageInViewport(
                builder,
                OfficeSvgImageRenderer.CreateDataUri(contentType, image.Bytes),
                projection,
                clipId,
                new OfficeImagePlacement(0D, 0D, snapshot.Width * scale, snapshot.Height * scale));
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
