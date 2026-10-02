using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private enum SvgFilterOperation { Blur, Offset, Matrix, Over, Blend }

    private sealed class SvgFilterNode {
        internal SvgFilterOperation Operation;
        internal int Input, Input2;
        internal double X, Y;
        internal double[]? Matrix;
        internal OfficeBlendMode BlendMode;
    }

    /// <summary>A resolved, acyclic list: input indices refer only to preceding results or the two source images.</summary>
    private sealed class SvgFilterGraph {
        internal readonly List<SvgFilterNode> Nodes = new();
        internal bool Linear = true, UserRegion;
        internal double ViewX, ViewY;
        internal string X = "-10%", Y = "-10%", Width = "120%", Height = "120%";
    }

    private static bool TryParseSvgFilterGraph(XElement filter, List<XElement> primitives, out SvgFilterGraph? graph) {
        graph = null;
        // Retain the established simple vector-filter contract. Routed and matrix/composite
        // operations use the bounded managed pixel graph instead of approximating their inputs.
        if (primitives.Count == 0 || primitives.Count > 32 || !primitives.Any(p =>
            p.Attribute("in") != null || p.Attribute("result") != null || p.Attribute("in2") != null ||
            p.Name.LocalName is "feColorMatrix" or "feComposite" or "feBlend")) return false;
        string units = filter.Attribute("filterUnits")?.Value ?? "objectBoundingBox";
        if (units is not ("objectBoundingBox" or "userSpaceOnUse") ||
            (filter.Attribute("primitiveUnits")?.Value ?? "userSpaceOnUse") != "userSpaceOnUse" ||
            filter.Attributes().Any(a => a.Name.LocalName is "href" or "filterRes")) return false;
        if (!TryFilterColorSpace(filter, out bool linear)) return false;
        var parsed = new SvgFilterGraph {
            Linear = linear, UserRegion = units == "userSpaceOnUse",
            X = filter.Attribute("x")?.Value ?? "-10%", Y = filter.Attribute("y")?.Value ?? "-10%",
            Width = filter.Attribute("width")?.Value ?? "120%", Height = filter.Attribute("height")?.Value ?? "120%"
        };
        var results = new Dictionary<string, int>(StringComparer.Ordinal) { ["SourceGraphic"] = -1, ["SourceAlpha"] = -2 };
        for (int index = 0; index < primitives.Count; index++) {
            XElement p = primitives[index];
            if (p.Name.Namespace != filter.Name.Namespace || p.HasElements ||
                p.Attributes().Any(a => a.Name.LocalName is "x" or "y" or "width" or "height" or "no-composite") ||
                !TryFilterColorSpace(p, out bool primitiveLinear) || primitiveLinear != linear ||
                !TryFilterInput(p.Attribute("in")?.Value, index, results, out int input)) return false;
            var node = new SvgFilterNode { Input = input };
            switch (p.Name.LocalName) {
                case "feGaussianBlur":
                    node.Operation = SvgFilterOperation.Blur;
                    if (!TryParseNumberList(p.Attribute("stdDeviation")?.Value ?? "0", 2, out IReadOnlyList<double> deviations) ||
                        deviations.Count is < 1 or > 2 || deviations.Any(d => !FiniteFilterNumber(d) || d < 0D || d > 64D) ||
                        (p.Attribute("edgeMode")?.Value ?? "none") != "none") return false;
                    node.X = deviations[0]; node.Y = deviations.Count == 2 ? deviations[1] : deviations[0];
                    break;
                case "feOffset":
                    node.Operation = SvgFilterOperation.Offset;
                    if (!TryParseFiniteNumber(p.Attribute("dx")?.Value, 0D, out node.X) ||
                        !TryParseFiniteNumber(p.Attribute("dy")?.Value, 0D, out node.Y)) return false;
                    break;
                case "feColorMatrix":
                    node.Operation = SvgFilterOperation.Matrix;
                    if ((p.Attribute("type")?.Value ?? "matrix") != "matrix") return false;
                    string values = p.Attribute("values")?.Value ?? "1 0 0 0 0 0 1 0 0 0 0 0 1 0 0 0 0 0 1 0";
                    if (!TryParseNumberList(values, 20, out IReadOnlyList<double> matrix) || matrix.Count != 20 ||
                        matrix.Any(d => !FiniteFilterNumber(d) || Math.Abs(d) > 10000D)) return false;
                    node.Matrix = matrix.ToArray();
                    break;
                case "feComposite":
                case "feBlend":
                    if (!TryFilterInput(p.Attribute("in2")?.Value, index, results, out node.Input2)) return false;
                    if (p.Name.LocalName == "feComposite") {
                        if ((p.Attribute("operator")?.Value ?? "over") != "over") return false;
                        node.Operation = SvgFilterOperation.Over;
                    } else {
                        node.Operation = SvgFilterOperation.Blend;
                        if (!TryParseBlendMode(p.Attribute("mode")?.Value ?? "normal", out node.BlendMode)) return false;
                        // These separable modes use the canonical shared blend component arithmetic.
                        if (node.BlendMode is OfficeBlendMode.Hue or OfficeBlendMode.Saturation or OfficeBlendMode.Color or OfficeBlendMode.Luminosity) return false;
                    }
                    break;
                default: return false;
            }
            parsed.Nodes.Add(node);
            string? name = p.Attribute("result")?.Value;
            if (name != null) {
                if (name.Length == 0 || name.Length > 256 || name.Any(char.IsWhiteSpace) || name is "SourceGraphic" or "SourceAlpha") return false;
                results[name] = index;
            }
        }
        graph = parsed;
        return true;
    }

    private static bool FiniteFilterNumber(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool TryFilterInput(string? input, int index, Dictionary<string, int> results, out int resolved) {
        resolved = index == 0 ? -1 : index - 1;
        return input == null || (input.Length <= 256 && results.TryGetValue(input, out resolved));
    }

    private static bool TryFilterColorSpace(XElement element, out bool linear) {
        linear = true;
        foreach (XElement ancestor in element.AncestorsAndSelf()) {
            string? value = ReadPresentationProperty(ancestor, "color-interpolation-filters")?.Trim();
            if (value == null || value is "inherit" or "unset") continue;
            if (value is "linearRGB" or "initial") return true;
            if (value is "sRGB" or "auto") { linear = false; return true; }
            return false;
        }
        return true;
    }
}
