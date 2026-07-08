using System;
using System.Linq;
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

            double x = ReadLength(element, "x", 0D, context, SvgLengthAxis.X);
            double y = ReadLength(element, "y", 0D, context, SvgLengthAxis.Y);
            double width = ReadLength(element, "width", image.Width, context, SvgLengthAxis.X);
            double height = ReadLength(element, "height", image.Height, context, SvgLengthAxis.Y);
            if (width <= 0D || height <= 0D) {
                return false;
            }

            double viewportX = x;
            double viewportY = y;
            double viewportWidth = width;
            double viewportHeight = height;
            bool slice = ApplyPreserveAspectRatio(element, image, ref x, ref y, ref width, ref height);
            OfficePoint topLeft = transform.Apply(x, y);
            OfficePoint topRight = transform.Apply(x + width, y);
            OfficePoint bottomLeft = transform.Apply(x, y + height);
            OfficePoint bottomRight = transform.Apply(x + width, y + height);
            OfficeRasterImage renderedImage = paint.Opacity < 1D ? ApplyImageOpacity(image, paint.Opacity) : image;
            using IDisposable? clipScope = slice
                ? canvas.PushClipPolygon(new[] {
                    transform.Apply(viewportX, viewportY),
                    transform.Apply(viewportX + viewportWidth, viewportY),
                    transform.Apply(viewportX + viewportWidth, viewportY + viewportHeight),
                    transform.Apply(viewportX, viewportY + viewportHeight)
                })
                : null;
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

        private static bool ApplyPreserveAspectRatio(XElement element, OfficeRasterImage image, ref double x, ref double y, ref double width, ref double height) {
            string raw = element.Attribute("preserveAspectRatio")?.Value?.Trim() ?? "xMidYMid meet";
            if (raw.Length == 0 || string.Equals(raw, "none", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            string[] parts = raw.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int alignIndex = parts.Length > 0 && string.Equals(parts[0], "defer", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
            string align = parts.Length > alignIndex ? parts[alignIndex] : "xMidYMid";
            bool slice = parts.Skip(alignIndex).Any(part => string.Equals(part, "slice", StringComparison.OrdinalIgnoreCase));
            double scaleX = width / image.Width;
            double scaleY = height / image.Height;
            double scale = slice ? Math.Max(scaleX, scaleY) : Math.Min(scaleX, scaleY);
            double scaledWidth = image.Width * scale;
            double scaledHeight = image.Height * scale;

            if (align.IndexOf("xMid", StringComparison.OrdinalIgnoreCase) >= 0) {
                x += (width - scaledWidth) / 2D;
            } else if (align.IndexOf("xMax", StringComparison.OrdinalIgnoreCase) >= 0) {
                x += width - scaledWidth;
            }

            if (align.IndexOf("YMid", StringComparison.OrdinalIgnoreCase) >= 0) {
                y += (height - scaledHeight) / 2D;
            } else if (align.IndexOf("YMax", StringComparison.OrdinalIgnoreCase) >= 0) {
                y += height - scaledHeight;
            }

            width = scaledWidth;
            height = scaledHeight;
            return slice;
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

            double determinant = (columnX * rowY) - (columnY * rowX);
            bool flipHorizontal = determinant < 0D;
            double placementX = topLeft.X;
            double placementY = topLeft.Y;
            double rotationColumnX = columnX;
            double rotationColumnY = columnY;
            if (flipHorizontal) {
                rotationColumnX = -rotationColumnX;
                rotationColumnY = -rotationColumnY;
                placementX += columnX;
                placementY += columnY;
            }

            double rotationDegrees = Math.Atan2(rotationColumnY, rotationColumnX) * 180D / Math.PI;
            projection = new OfficeImageProjection(
                new OfficeImagePlacement(placementX, placementY, width, height),
                rotationDegrees: rotationDegrees,
                rotationCenterX: placementX,
                rotationCenterY: placementY,
                flipHorizontal: flipHorizontal);
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
