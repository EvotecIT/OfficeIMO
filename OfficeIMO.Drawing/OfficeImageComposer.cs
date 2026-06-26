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
        var svgLayers = new List<OfficeImageLayer>();
        var svgLayerIds = new List<HashSet<string>>();
        var seenIds = new HashSet<string>(StringComparer.Ordinal);
        var duplicateIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (OfficeImageLayer layer in layers) {
            if (layer.SvgInnerContent != null) {
                HashSet<string> ids = GetSvgIds(layer.SvgInnerContent);
                foreach (string id in ids) {
                    if (!seenIds.Add(id)) {
                        duplicateIds.Add(id);
                    }
                }

                svgLayers.Add(layer);
                svgLayerIds.Add(ids);
            }
        }

        for (int index = 0; index < svgLayers.Count; index++) {
            OfficeImageLayer layer = svgLayers[index];
            string layerContent = layer.SvgInnerContent!;
            if (duplicateIds.Count > 0 && svgLayerIds[index].Overlaps(duplicateIds)) {
                layerContent = NamespaceSvgLayerIds(layerContent, "officeimo-layer-" + (index + 1) + "-", duplicateIds);
            }

            builder.AppendNestedSvg(layer.X, layer.Y, layer.Width, layer.Height, layerContent);
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

    private static string NamespaceSvgLayerIds(string svg, string prefix, HashSet<string> idsToNamespace) {
        var ids = new Dictionary<string, string>(StringComparer.Ordinal);
        CollectSvgIds(svg, '"', prefix, idsToNamespace, ids);
        CollectSvgIds(svg, '\'', prefix, idsToNamespace, ids);
        if (ids.Count == 0) {
            return svg;
        }

        string namespaced = svg;
        foreach (KeyValuePair<string, string> id in ids) {
            namespaced = ReplaceOrdinal(namespaced, "id=\"" + id.Key + "\"", "id=\"" + id.Value + "\"");
            namespaced = ReplaceOrdinal(namespaced, "id='" + id.Key + "'", "id='" + id.Value + "'");
        }

        foreach (KeyValuePair<string, string> id in ids) {
            namespaced = ReplaceOrdinal(namespaced, "url(#" + id.Key + ")", "url(#" + id.Value + ")");
            namespaced = ReplaceOrdinal(namespaced, "\"#" + id.Key + "\"", "\"#" + id.Value + "\"");
            namespaced = ReplaceOrdinal(namespaced, "'#" + id.Key + "'", "'#" + id.Value + "'");
        }

        return namespaced;
    }

    private static HashSet<string> GetSvgIds(string svg) {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        CollectSvgIds(svg, '"', ids);
        CollectSvgIds(svg, '\'', ids);
        return ids;
    }

    private static void CollectSvgIds(string svg, char quote, HashSet<string> ids) {
        string marker = "id=" + quote;
        int searchStart = 0;
        while (searchStart < svg.Length) {
            int markerIndex = svg.IndexOf(marker, searchStart, StringComparison.Ordinal);
            if (markerIndex < 0) {
                return;
            }

            searchStart = markerIndex + marker.Length;
            if (markerIndex > 0 && IsSvgNameChar(svg[markerIndex - 1])) {
                continue;
            }

            int valueStart = markerIndex + marker.Length;
            int valueEnd = svg.IndexOf(quote, valueStart);
            if (valueEnd < 0) {
                return;
            }

            string id = svg.Substring(valueStart, valueEnd - valueStart);
            if (id.Length > 0) {
                ids.Add(id);
            }

            searchStart = valueEnd + 1;
        }
    }

    private static void CollectSvgIds(string svg, char quote, string prefix, HashSet<string> idsToNamespace, Dictionary<string, string> ids) {
        string marker = "id=" + quote;
        int searchStart = 0;
        while (searchStart < svg.Length) {
            int markerIndex = svg.IndexOf(marker, searchStart, StringComparison.Ordinal);
            if (markerIndex < 0) {
                return;
            }

            searchStart = markerIndex + marker.Length;
            if (markerIndex > 0 && IsSvgNameChar(svg[markerIndex - 1])) {
                continue;
            }

            int valueStart = markerIndex + marker.Length;
            int valueEnd = svg.IndexOf(quote, valueStart);
            if (valueEnd < 0) {
                return;
            }

            string id = svg.Substring(valueStart, valueEnd - valueStart);
            if (id.Length > 0 && idsToNamespace.Contains(id) && !ids.ContainsKey(id)) {
                ids.Add(id, prefix + id);
            }

            searchStart = valueEnd + 1;
        }
    }

    private static bool IsSvgNameChar(char value) =>
        (value >= 'A' && value <= 'Z') ||
        (value >= 'a' && value <= 'z') ||
        (value >= '0' && value <= '9') ||
        value == '_' ||
        value == '-' ||
        value == ':';

    private static string ReplaceOrdinal(string value, string oldValue, string newValue) {
        int index = value.IndexOf(oldValue, StringComparison.Ordinal);
        if (index < 0) {
            return value;
        }

        var builder = new StringBuilder(value.Length + newValue.Length - oldValue.Length);
        int start = 0;
        while (index >= 0) {
            builder.Append(value, start, index - start);
            builder.Append(newValue);
            start = index + oldValue.Length;
            index = value.IndexOf(oldValue, start, StringComparison.Ordinal);
        }

        builder.Append(value, start, value.Length - start);
        return builder.ToString();
    }
}
