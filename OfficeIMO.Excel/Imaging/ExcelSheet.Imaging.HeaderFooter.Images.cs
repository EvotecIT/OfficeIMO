using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const double HeaderFooterImagePadding = 4D;

        private static HeaderFooterImageSnapshot? SelectHeaderFooterImage(HeaderFooterImageSnapshot? image, string sectionText) =>
            image != null && sectionText.IndexOf("&G", StringComparison.Ordinal) >= 0 ? image : null;

        private static int ResolveHeaderFooterBandHeight(double imageHeightPoints, double scale) {
            double textHeight = HeaderFooterBandHeight * scale;
            double imageHeight = imageHeightPoints <= 0D
                ? 0D
                : PointsToPixels(imageHeightPoints, scale) + (HeaderFooterImagePadding * 2D * scale);
            return Math.Max(1, (int)Math.Ceiling(Math.Max(textHeight, imageHeight)));
        }

        private static void AddHeaderFooterImageDiagnostics(
            OfficeImageExportFormat format,
            HeaderFooterTextChrome chrome,
            string source,
            List<OfficeImageExportDiagnostic> diagnostics) {
            int imageCount = 0;
            int unsupportedCount = 0;
            AddHeaderFooterImageDiagnostics(format, chrome.HeaderLeftImage, source, "header left", diagnostics, ref imageCount, ref unsupportedCount);
            AddHeaderFooterImageDiagnostics(format, chrome.HeaderCenterImage, source, "header center", diagnostics, ref imageCount, ref unsupportedCount);
            AddHeaderFooterImageDiagnostics(format, chrome.HeaderRightImage, source, "header right", diagnostics, ref imageCount, ref unsupportedCount);
            AddHeaderFooterImageDiagnostics(format, chrome.FooterLeftImage, source, "footer left", diagnostics, ref imageCount, ref unsupportedCount);
            AddHeaderFooterImageDiagnostics(format, chrome.FooterCenterImage, source, "footer center", diagnostics, ref imageCount, ref unsupportedCount);
            AddHeaderFooterImageDiagnostics(format, chrome.FooterRightImage, source, "footer right", diagnostics, ref imageCount, ref unsupportedCount);
            if (imageCount > unsupportedCount) {
                diagnostics.Add(new OfficeImageExportDiagnostic(
                    OfficeImageExportDiagnosticSeverity.Info,
                    ExcelImageExportDiagnosticCodes.HeaderFooterImageApproximation,
                    "Worksheet header/footer images were rendered through the dependency-free image approximation path.",
                    source));
            }
        }

        private static void AddHeaderFooterImageDiagnostics(
            OfficeImageExportFormat format,
            HeaderFooterImageSnapshot? image,
            string source,
            string location,
            List<OfficeImageExportDiagnostic> diagnostics,
            ref int imageCount,
            ref int unsupportedCount) {
            if (image == null) {
                return;
            }

            imageCount++;
            if (CanRenderHeaderFooterImage(format, image)) {
                return;
            }

            unsupportedCount++;
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.HeaderFooterImageUnsupported,
                "Worksheet " + location + " image with content type '" + image.ContentType + "' cannot be rendered by " + format + " header/footer image export yet.",
                source));
        }

        private static bool CanRenderHeaderFooterImage(OfficeImageExportFormat format, HeaderFooterImageSnapshot image) {
            if (format == OfficeImageExportFormat.Svg) {
                return OfficeSvgImageRenderer.TryResolveEmbeddableContentType(image.ContentType, image.Bytes, null, out _);
            }

            return OfficePngReader.TryDecode(image.Bytes, out _);
        }

        private static void DrawHeaderFooterRasterImages(
            OfficeRasterCanvas canvas,
            HeaderFooterTextChrome chrome,
            bool isHeader,
            double bandTop,
            double bandHeight,
            OfficeTextZoneLayout zones,
            double scale) {
            DrawHeaderFooterRasterImage(canvas, isHeader ? chrome.HeaderLeftImage : chrome.FooterLeftImage, zones.Left, bandTop, bandHeight, scale, OfficeTextAlignment.Left);
            DrawHeaderFooterRasterImage(canvas, isHeader ? chrome.HeaderCenterImage : chrome.FooterCenterImage, zones.Center, bandTop, bandHeight, scale, OfficeTextAlignment.Center);
            DrawHeaderFooterRasterImage(canvas, isHeader ? chrome.HeaderRightImage : chrome.FooterRightImage, zones.Right, bandTop, bandHeight, scale, OfficeTextAlignment.Right);
        }

        private static void DrawHeaderFooterRasterImage(
            OfficeRasterCanvas canvas,
            HeaderFooterImageSnapshot? image,
            OfficeTextZone zone,
            double bandTop,
            double bandHeight,
            double scale,
            OfficeTextAlignment alignment) {
            if (image == null || !OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster) || raster == null) {
                return;
            }

            (double x, double y, double width, double height) = ResolveHeaderFooterImageBox(image, zone, bandTop, bandHeight, scale, alignment);
            using (canvas.PushClipRectangle(zone.X, bandTop, zone.Width, bandHeight)) {
                canvas.DrawImage(raster, x, y, width, height);
            }
        }

        private static void AppendHeaderFooterSvgImages(
            StringBuilder builder,
            HeaderFooterTextChrome chrome,
            bool isHeader,
            double bandTop,
            double bandHeight,
            OfficeTextZoneLayout zones,
            double scale) {
            AppendHeaderFooterSvgImage(builder, isHeader ? chrome.HeaderLeftImage : chrome.FooterLeftImage, zones.Left, bandTop, bandHeight, scale, OfficeTextAlignment.Left, isHeader ? "header-left-image" : "footer-left-image");
            AppendHeaderFooterSvgImage(builder, isHeader ? chrome.HeaderCenterImage : chrome.FooterCenterImage, zones.Center, bandTop, bandHeight, scale, OfficeTextAlignment.Center, isHeader ? "header-center-image" : "footer-center-image");
            AppendHeaderFooterSvgImage(builder, isHeader ? chrome.HeaderRightImage : chrome.FooterRightImage, zones.Right, bandTop, bandHeight, scale, OfficeTextAlignment.Right, isHeader ? "header-right-image" : "footer-right-image");
        }

        private static void AppendHeaderFooterSvgImage(
            StringBuilder builder,
            HeaderFooterImageSnapshot? image,
            OfficeTextZone zone,
            double bandTop,
            double bandHeight,
            double scale,
            OfficeTextAlignment alignment,
            string clipSuffix) {
            if (image == null || !OfficeSvgImageRenderer.TryResolveEmbeddableContentType(image.ContentType, image.Bytes, null, out string contentType)) {
                return;
            }

            (double x, double y, double width, double height) = ResolveHeaderFooterImageBox(image, zone, bandTop, bandHeight, scale, alignment);
            string clipId = "xl-header-footer-" + clipSuffix;
            OfficeSvgImageRenderer.AppendImageInViewport(
                builder,
                OfficeSvgImageRenderer.CreateDataUri(contentType, image.Bytes),
                new OfficeImageProjection(new OfficeImagePlacement(x, y, width, height)),
                clipId,
                new OfficeImagePlacement(zone.X, bandTop, zone.Width, bandHeight));
        }

        private static (double X, double Y, double Width, double Height) ResolveHeaderFooterImageBox(
            HeaderFooterImageSnapshot image,
            OfficeTextZone zone,
            double bandTop,
            double bandHeight,
            double scale,
            OfficeTextAlignment alignment) {
            double width = PointsToPixels(image.WidthPoints, scale);
            double height = PointsToPixels(image.HeightPoints, scale);
            double x = zone.X;
            if (alignment == OfficeTextAlignment.Center) {
                x = zone.AnchorX - (width / 2D);
            } else if (alignment == OfficeTextAlignment.Right) {
                x = zone.AnchorX - width;
            }

            double y = bandTop + Math.Max(0D, (bandHeight - height) / 2D);
            return (x, y, width, height);
        }

        private static double PointsToPixels(double points, double scale) =>
            Math.Max(1D, points * ImageExportDpi / 72D * scale);
    }
}
