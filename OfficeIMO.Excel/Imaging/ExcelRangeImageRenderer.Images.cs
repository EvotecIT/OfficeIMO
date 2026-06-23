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

            double x = image.X * scale;
            double y = image.Y * scale;
            double width = image.Width * scale;
            double height = image.Height * scale;
            canvas.DrawImage(
                raster,
                x,
                y,
                width,
                height,
                image.CropLeftRatio,
                image.CropTopRatio,
                GetVisibleCropWidth(image),
                GetVisibleCropHeight(image),
                image.RotationDegrees,
                x + (width / 2D),
                y + (height / 2D),
                image.FlipHorizontal,
                image.FlipVertical);
        }

        private static void AppendSvgImage(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelVisualImage image, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics, ref int index) {
            double scale = options.Scale;
            if (!TryResolveSvgImageContentType(image, out string contentType)) {
                diagnostics?.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ImageSvgFormatUnsupported,
                    "Worksheet image format '" + image.DetectedFormat + "' cannot be embedded reliably in SVG output by the dependency-free image renderer yet.",
                    image.Source));
                return;
            }

            string clipId = "xl-image-clip-" + (++index).ToString(System.Globalization.CultureInfo.InvariantCulture);
            double x = image.X * scale;
            double y = image.Y * scale;
            double width = image.Width * scale;
            double height = image.Height * scale;
            double visibleWidth = GetVisibleCropWidth(image);
            double visibleHeight = GetVisibleCropHeight(image);

            OfficeSvgImageRenderer.AppendImage(
                builder,
                OfficeSvgImageRenderer.CreateDataUri(contentType, image.Bytes),
                x,
                y,
                width,
                height,
                clipId,
                image.HasCrop ? x : 0D,
                image.HasCrop ? y : 0D,
                image.HasCrop ? width : snapshot.Width * scale,
                image.HasCrop ? height : snapshot.Height * scale,
                image.CropLeftRatio,
                image.CropTopRatio,
                visibleWidth,
                visibleHeight,
                image.RotationDegrees,
                image.FlipHorizontal,
                image.FlipVertical);
        }

        private static bool TryResolveSvgImageContentType(ExcelVisualImage image, out string contentType) {
            switch (image.DetectedFormat) {
                case OfficeImageFormat.Png:
                    contentType = "image/png";
                    return true;
                case OfficeImageFormat.Jpeg:
                    contentType = "image/jpeg";
                    return true;
                case OfficeImageFormat.Gif:
                    contentType = "image/gif";
                    return true;
                case OfficeImageFormat.Svg:
                    contentType = "image/svg+xml";
                    return true;
                default:
                    contentType = string.Empty;
                    return false;
            }
        }

        private static double GetVisibleCropWidth(ExcelVisualImage image) =>
            Math.Max(0.001D, 1D - image.CropLeftRatio - image.CropRightRatio);

        private static double GetVisibleCropHeight(ExcelVisualImage image) =>
            Math.Max(0.001D, 1D - image.CropTopRatio - image.CropBottomRatio);
    }
}
