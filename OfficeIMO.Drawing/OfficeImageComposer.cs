using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free image page composer for raster and SVG layer output.
/// </summary>
public static class OfficeImageComposer {
    /// <summary>
    /// Composes raster layers onto a new image and returns encoded PNG bytes.
    /// </summary>
    public static byte[] ComposePng(
        int width,
        int height,
        OfficeColor backgroundColor,
        IEnumerable<OfficeImageLayer> layers,
        Action<OfficeRasterCanvas>? beforeLayers = null,
        Action<OfficeRasterCanvas>? afterLayers = null) =>
        OfficePngWriter.Encode(ComposeRaster(width, height, backgroundColor, layers, beforeLayers, afterLayers));

    /// <summary>
    /// Composes raster layers onto a new image.
    /// </summary>
    public static OfficeRasterImage ComposeRaster(
        int width,
        int height,
        OfficeColor backgroundColor,
        IEnumerable<OfficeImageLayer> layers,
        Action<OfficeRasterCanvas>? beforeLayers = null,
        Action<OfficeRasterCanvas>? afterLayers = null) {
        ValidateOutputSize(width, height);
        if (layers == null) {
            throw new ArgumentNullException(nameof(layers));
        }

        OfficeRasterImage image = new OfficeRasterImage(width, height, backgroundColor);
        var canvas = new OfficeRasterCanvas(image);
        beforeLayers?.Invoke(canvas);
        foreach (OfficeImageLayer layer in layers) {
            if (layer.RasterImage != null) {
                canvas.DrawImage(layer.RasterImage, layer.X, layer.Y, layer.Width, layer.Height);
            }
        }

        afterLayers?.Invoke(canvas);
        return image;
    }

    /// <summary>
    /// Composes SVG layers into a root SVG document and returns UTF-8 encoded bytes.
    /// </summary>
    public static byte[] ComposeSvgBytes(
        int width,
        int height,
        OfficeColor backgroundColor,
        IEnumerable<OfficeImageLayer> layers,
        Action<StringBuilder>? beforeLayers = null,
        Action<StringBuilder>? afterLayers = null) =>
        Encoding.UTF8.GetBytes(ComposeSvg(width, height, backgroundColor, layers, beforeLayers, afterLayers));

    /// <summary>
    /// Composes SVG layers into a root SVG document.
    /// </summary>
    public static string ComposeSvg(
        int width,
        int height,
        OfficeColor backgroundColor,
        IEnumerable<OfficeImageLayer> layers,
        Action<StringBuilder>? beforeLayers = null,
        Action<StringBuilder>? afterLayers = null) {
        ValidateOutputSize(width, height);
        if (layers == null) {
            throw new ArgumentNullException(nameof(layers));
        }

        var builder = new StringBuilder();
        builder.Append("<svg xmlns=\"http://www.w3.org/2000/svg\"")
            .AppendNumberAttribute("width", width)
            .AppendNumberAttribute("height", height)
            .AppendAttribute("viewBox", "0 0 " + OfficeSvgFormatting.FormatNumber(width) + " " + OfficeSvgFormatting.FormatNumber(height))
            .Append('>');

        var backgroundAttributes = new StringBuilder();
        backgroundAttributes.AppendPaintAttribute("fill", backgroundColor);
        builder.AppendRectElement(0D, 0D, width, height, backgroundAttributes.ToString());
        beforeLayers?.Invoke(builder);
        foreach (OfficeImageLayer layer in layers) {
            if (layer.SvgInnerContent != null) {
                builder.AppendNestedSvg(layer.X, layer.Y, layer.Width, layer.Height, layer.SvgInnerContent);
            }
        }

        afterLayers?.Invoke(builder);
        builder.Append("</svg>");
        return builder.ToString();
    }

    private static void ValidateOutputSize(int width, int height) {
        if (width <= 0) {
            throw new ArgumentOutOfRangeException(nameof(width), "Output width must be positive.");
        }

        if (height <= 0) {
            throw new ArgumentOutOfRangeException(nameof(height), "Output height must be positive.");
        }
    }
}
