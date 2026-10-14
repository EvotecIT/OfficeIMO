using OfficeIMO.Drawing;
using System.Text.RegularExpressions;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    // Canvas dimensions stay positive for composition. Semantic frame dimensions
    // may be zero: frame paint is omitted and nonempty text reports clipping.
    private readonly struct TextFrameProjection {
        internal TextFrameProjection(OfficeDrawing drawing, double width, double height) {
            Drawing = drawing; Width = width; Height = height;
        }
        internal OfficeDrawing Drawing { get; }
        internal double Width { get; }
        internal double Height { get; }
    }

    private sealed class TextBoxConstraints {
        internal TextBoxConstraints(double? maximumWidth, double? maximumHeight) {
            MaximumWidth = maximumWidth; MaximumHeight = maximumHeight;
        }
        internal double? MaximumWidth { get; }
        internal double? MaximumHeight { get; }
    }

    private static bool IsOrdinaryTextBox(OdgShape shape) => shape.ElementName == "frame" && !shape.IsImage &&
        shape.TextRoot.Name == OdfNamespaces.Draw + "text-box";

    private static TextBoxConstraints? ResolveTextBoxConstraints(OdgShape shape, ref double width, ref double height) {
        if (!IsOrdinaryTextBox(shape) || shape.ReadGraphicProperty(OdfNamespaces.Draw + "auto-grow-width") is not ("false" or "true") ||
            shape.ReadGraphicProperty(OdfNamespaces.Draw + "auto-grow-height") is not ("false" or "true") ||
            shape.Element.Attribute(OdfNamespaces.Style + "rel-width") != null ||
            shape.Element.Attribute(OdfNamespaces.Style + "rel-height") != null ||
            (string?)shape.Element.Attribute(OdfNamespaces.Text + "anchor-type") is not (null or "page")) return null;

        if (!ReadPair("width", out double? minWidth, out double? maxWidth) ||
            !ReadPair("height", out double? minHeight, out double? maxHeight)) return null;
        // ODF 1.4 sections 19.240/19.241: instance minima override SVG frame
        // dimensions. Graphic-style creation defaults do not resize saved frames.
        width = minWidth ?? width; height = minHeight ?? height;
        return new TextBoxConstraints(maxWidth, maxHeight);

        bool ReadPair(string axis, out double? minimum, out double? maximum) {
            string? minRaw = (string?)shape.TextRoot.Attribute(OdfNamespaces.Fo + "min-" + axis);
            string? maxRaw = (string?)shape.TextRoot.Attribute(OdfNamespaces.Fo + "max-" + axis);
            minimum = maximum = null;
            if (!ReadLength(minRaw, out minimum) || !ReadLength(maxRaw, out maximum)) return false;
            // A maximum is defined with a corresponding minimum, in matching
            // lexical units. Contradictory bounds remain preserved, unsupported.
            return maxRaw == null || minRaw != null &&
                minRaw.Substring(minRaw.Length - 2) == maxRaw.Substring(maxRaw.Length - 2) && maximum >= minimum;
        }
    }

    private static bool ReadLength(string? raw, out double? points) {
        points = null;
        if (raw == null) return true;
        if (!Regex.IsMatch(raw, @"\A(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:cm|mm|in|pt|pc)\z", RegexOptions.CultureInvariant) ||
            !OdfLength.Parse(raw).TryToPoints(out double value)) return false;
        points = value; return true;
    }

    private static OfficeDrawing ResizeTextCanvas(OfficeDrawing target, double width, double height) {
        double canvasWidth = Math.Max(width, .001D), canvasHeight = Math.Max(height, .001D);
        if (target.Width == canvasWidth && target.Height == canvasHeight) return target;
        var resized = new OfficeDrawing(canvasWidth, canvasHeight);
        CopyDrawingResources(target, resized);
        return resized;
    }
}
