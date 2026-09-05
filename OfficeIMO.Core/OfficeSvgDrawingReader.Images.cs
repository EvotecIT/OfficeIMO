using System;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryAddEmbeddedSvgImage(
        XElement element,
        OfficeDrawing drawing,
        SvgPaintContext style,
        OfficeTransform transform,
        double viewX,
        double viewY) {
        XAttribute[] hrefAttributes = element.Attributes()
            .Where(attribute => attribute.Name.LocalName.Equals("href", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (hrefAttributes.Length != 1
            || !TryDecodeEmbeddedRasterImage(hrefAttributes[0].Value, out byte[] bytes, out string contentType, out OfficeImageInfo info)
            || !TryViewportLength(element, "width", drawing.Width, out double width)
            || !TryViewportLength(element, "height", drawing.Height, out double height)
            || width <= 0D
            || height <= 0D) return false;

        double x = ReadViewportCoordinate(element, "x", viewX, drawing.Width);
        double y = ReadViewportCoordinate(element, "y", viewY, drawing.Height);
        if (!TryParsePreserveAspectRatio(element.Attribute("preserveAspectRatio")?.Value,
                out SvgAspectAlignment alignment, out bool slice)) return false;
        OfficeImageProjection projection = ResolveSvgImageProjection(
            x, y, width, height, info.Width, info.Height, alignment, slice);
        string? alternativeText = element.Attributes()
            .FirstOrDefault(attribute => attribute.Name.LocalName.Equals("aria-label", StringComparison.OrdinalIgnoreCase))?.Value;

        var imageLayer = new OfficeDrawing(drawing.Width, drawing.Height);
        imageLayer.AddClippedImage(
            bytes,
            contentType,
            projection,
            0D,
            0D,
            OfficeClipPath.Rectangle(drawing.Width, drawing.Height),
            alternativeText,
            style.Opacity);
        drawing.AddEffectDrawing(imageLayer, transform);
        return true;
    }

    private static bool TryDecodeEmbeddedRasterImage(
        string? href,
        out byte[] bytes,
        out string contentType,
        out OfficeImageInfo info) {
        bytes = Array.Empty<byte>();
        contentType = string.Empty;
        info = null!;
        if (string.IsNullOrWhiteSpace(href) || !href!.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return false;
        int comma = href.IndexOf(',');
        if (comma <= 5 || comma == href.Length - 1) return false;
        string metadata = href.Substring(5, comma - 5);
        int separator = metadata.IndexOf(';');
        contentType = (separator < 0 ? metadata : metadata.Substring(0, separator)).Trim();
        if (!contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
            || !ContainsBase64DataUriToken(metadata)) return false;
        try {
            bytes = Convert.FromBase64String(href.Substring(comma + 1));
        } catch (FormatException) {
            return false;
        }
        return OfficeImageReader.TryIdentifyByContent(bytes, null, out info)
            && info.Format != OfficeImageFormat.Svg
            && OfficeRasterGuards.TryEnsurePixelCount(info.Width, info.Height, out _);
    }

    private static OfficeImageProjection ResolveSvgImageProjection(
        double x,
        double y,
        double width,
        double height,
        double intrinsicWidth,
        double intrinsicHeight,
        SvgAspectAlignment alignment,
        bool slice) {
        if (alignment == SvgAspectAlignment.None || intrinsicWidth <= 0D || intrinsicHeight <= 0D) {
            return new OfficeImageProjection(new OfficeImagePlacement(x, y, width, height));
        }
        ResolveAlignmentFactors(alignment, out double alignX, out double alignY);
        double scale = slice
            ? Math.Max(width / intrinsicWidth, height / intrinsicHeight)
            : Math.Min(width / intrinsicWidth, height / intrinsicHeight);
        double renderedWidth = intrinsicWidth * scale;
        double renderedHeight = intrinsicHeight * scale;
        if (!slice) {
            return new OfficeImageProjection(new OfficeImagePlacement(
                x + (width - renderedWidth) * alignX,
                y + (height - renderedHeight) * alignY,
                renderedWidth,
                renderedHeight));
        }
        double visibleWidth = Math.Min(1D, width / renderedWidth);
        double visibleHeight = Math.Min(1D, height / renderedHeight);
        double horizontalCrop = 1D - visibleWidth;
        double verticalCrop = 1D - visibleHeight;
        OfficeImageSourceCrop crop = OfficeImageSourceCrop.FromStrictFractions(
            horizontalCrop * alignX,
            verticalCrop * alignY,
            horizontalCrop * (1D - alignX),
            verticalCrop * (1D - alignY));
        return new OfficeImageProjection(new OfficeImagePlacement(x, y, width, height), crop);
    }
}
