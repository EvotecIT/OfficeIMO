using System;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static bool RenderImage(OfficeRasterCanvas canvas, XElement element, SvgPaint paint, SvgTransform transform, SvgRenderContext context) {
            if (paint.Opacity <= 0D) {
                return false;
            }

            string? href = ReadHref(element);
            if (!TryReadImageBytes(href, context, out byte[]? bytes) ||
                !OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? image) ||
                image == null) {
                return false;
            }

            double x = ReadLength(element, "x", 0D);
            double y = ReadLength(element, "y", 0D);
            double width = ReadLength(element, "width", image.Width);
            double height = ReadLength(element, "height", image.Height);
            if (width <= 0D || height <= 0D) {
                return false;
            }

            OfficePoint topLeft = transform.Apply(x, y);
            OfficePoint topRight = transform.Apply(x + width, y);
            OfficePoint bottomLeft = transform.Apply(x, y + height);
            OfficePoint bottomRight = transform.Apply(x + width, y + height);
            OfficeRasterImage renderedImage = paint.Opacity < 1D ? ApplyImageOpacity(image, paint.Opacity) : image;
            if (TryCreateImageProjection(topLeft, topRight, bottomLeft, bottomRight, out OfficeImageProjection projection)) {
                canvas.DrawImage(renderedImage, projection);
                return true;
            }

            double left = Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomLeft.X, bottomRight.X));
            double top = Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomLeft.Y, bottomRight.Y));
            double right = Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomLeft.X, bottomRight.X));
            double bottom = Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomLeft.Y, bottomRight.Y));
            if (right <= left || bottom <= top) {
                return false;
            }

            canvas.DrawImage(renderedImage, left, top, right - left, bottom - top);
            return true;
        }

        private static bool TryCreateImageProjection(OfficePoint topLeft, OfficePoint topRight, OfficePoint bottomLeft, OfficePoint bottomRight, out OfficeImageProjection projection) {
            projection = default;
            double columnX = topRight.X - topLeft.X;
            double columnY = topRight.Y - topLeft.Y;
            double rowX = bottomLeft.X - topLeft.X;
            double rowY = bottomLeft.Y - topLeft.Y;
            double width = Math.Sqrt(columnX * columnX + columnY * columnY);
            double height = Math.Sqrt(rowX * rowX + rowY * rowY);
            if (width <= 0D || height <= 0D) {
                return false;
            }

            double dot = columnX * rowX + columnY * rowY;
            if (Math.Abs(dot) > 0.001D * width * height) {
                return false;
            }

            if (Math.Abs(bottomRight.X - (topRight.X + rowX)) > 0.001D ||
                Math.Abs(bottomRight.Y - (topRight.Y + rowY)) > 0.001D) {
                return false;
            }

            double rotationDegrees = Math.Atan2(columnY, columnX) * 180D / Math.PI;
            projection = new OfficeImageProjection(
                new OfficeImagePlacement(topLeft.X, topLeft.Y, width, height),
                rotationDegrees: rotationDegrees,
                rotationCenterX: topLeft.X,
                rotationCenterY: topLeft.Y);
            return true;
        }

        private static OfficeRasterImage ApplyImageOpacity(OfficeRasterImage image, double opacity) {
            OfficeRasterImage adjusted = new(image.Width, image.Height);
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    adjusted.SetPixel(x, y, OfficeColor.FromRgba(pixel.R, pixel.G, pixel.B, (byte)Math.Round(pixel.A * opacity)));
                }
            }

            return adjusted;
        }

        private static bool TryReadImageBytes(string? href, SvgRenderContext context, out byte[]? bytes) {
            if (TryReadDataUriBytes(href, out bytes)) {
                return true;
            }

            bytes = null;
            if (string.IsNullOrWhiteSpace(href)) {
                return false;
            }

            return context.TryGetImageBytes(href!.Trim(), out bytes);
        }

        private static bool TryReadDataUriBytes(string? href, out byte[]? bytes) {
            bytes = null;
            if (string.IsNullOrWhiteSpace(href)) {
                return false;
            }

            string value = href!.Trim();
            if (!value.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            int comma = value.IndexOf(',');
            if (comma < 0) {
                return false;
            }

            string metadata = value.Substring(5, comma - 5);
            string payload = value.Substring(comma + 1);
            if (metadata.IndexOf(";base64", StringComparison.OrdinalIgnoreCase) < 0 ||
                payload.Length == 0) {
                return false;
            }

            try {
                bytes = Convert.FromBase64String(payload);
                return true;
            } catch (FormatException) {
                return false;
            }
        }
    }
}
