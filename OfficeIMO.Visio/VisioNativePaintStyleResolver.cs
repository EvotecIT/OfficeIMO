using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    /// <summary>Resolves cached native paint properties into the loaded model without changing source XML or evaluating formulas.</summary>
    internal sealed class VisioNativePaintStyleResolver {
        private readonly Dictionary<string, XElement> _styles = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _colors = new(StringComparer.Ordinal);
        private readonly Dictionary<string, Dictionary<string, string>> _cache = new(StringComparer.Ordinal);

        internal VisioNativePaintStyleResolver(XElement? document) {
            if (document == null) return;
            XNamespace ns = document.Name.Namespace;
            foreach (var group in document.Element(ns + "StyleSheets")?.Elements(ns + "StyleSheet")
                .Where(style => style.Attribute("ID") != null).GroupBy(style => (string)style.Attribute("ID")!, StringComparer.Ordinal)
                ?? Enumerable.Empty<IGrouping<string, XElement>>()) {
                if (group.Count() == 1) _styles.Add(group.Key, group.Single());
            }
            foreach (XElement color in document.Element(ns + "Colors")?.Elements(ns + "ColorEntry") ?? Enumerable.Empty<XElement>()) {
                string? index = (string?)color.Attribute("IX"), rgb = (string?)color.Attribute("RGB");
                if (index != null && rgb != null && !_colors.ContainsKey(index)) _colors.Add(index, rgb);
            }
        }

        internal void Apply(VisioShape shape, XElement source) {
            var cells = Cells(source, "LineStyle", "FillStyle");
            if (Number(cells, "LineWeight", out double weight) && weight >= 0) shape.LineWeight = weight;
            if (Number(cells, "LinePattern", out double linePattern) && linePattern >= 0 && linePattern <= 255) shape.LinePattern = (int)linePattern;
            if (Number(cells, "FillPattern", out double fillPattern) && fillPattern >= 0 && fillPattern <= 255) shape.FillPattern = (int)fillPattern;
            shape.LineColor = Transparency(cells, "LineColorTrans", Color(cells, "LineColor", shape.LineColor));
            shape.FillColor = Transparency(cells, "FillForegndTrans", Color(cells, "FillForegnd", shape.FillColor));
        }

        internal void Apply(VisioConnector connector, XElement source, VisioShape? inherited = null) {
            if (inherited != null) {
                connector.LineWeight = inherited.LineWeight;
                connector.LinePattern = inherited.LinePattern;
                connector.LineColor = inherited.LineColor;
            }
            var cells = Cells(source, "LineStyle");
            if (Number(cells, "LineWeight", out double weight) && weight >= 0) connector.LineWeight = weight;
            if (Number(cells, "LinePattern", out double pattern) && pattern >= 0 && pattern <= 255) connector.LinePattern = (int)pattern;
            OfficeColor color = Color(cells, "LineColor", connector.LineColor);
            if (inherited != null) color = InheritTransparency(color, inherited.LineColor);
            connector.LineColor = Transparency(cells, "LineColorTrans", color);
        }

        private Dictionary<string, string> Cells(XElement source, params string[] bindings) {
            var cells = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string binding in bindings) {
                foreach (var cell in Style((string?)source.Attribute(binding), binding)) cells[cell.Key] = cell.Value;
            }
            foreach (XElement cell in source.Elements(source.Name.Namespace + "Cell")) {
                string? name = (string?)cell.Attribute("N"), value = (string?)cell.Attribute("V");
                if (name != null && value != null) cells[name] = value;
            }
            return cells;
        }

        private Dictionary<string, string> Style(string? reference, string binding) {
            string key = binding + ":" + reference;
            if (_cache.TryGetValue(key, out var cached)) return cached;
            var cells = new Dictionary<string, string>(StringComparer.Ordinal);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            string prefix = binding == "LineStyle" ? "Line" : "Fill";
            string enabled = binding == "LineStyle" ? "EnableLineProps" : "EnableFillProps";
            // Closest concrete caches win, including F=Inh with a native V cache.
            while (reference != null && seen.Count < 64 && seen.Add(reference) && _styles.TryGetValue(reference, out XElement? style)) {
                XNamespace ns = style.Name.Namespace;
                if (style.Elements(ns + "Cell").Any(cell => (string?)cell.Attribute("N") == enabled && (string?)cell.Attribute("V") == "0")) break;
                foreach (XElement cell in style.Elements(ns + "Cell")) {
                    string? name = (string?)cell.Attribute("N"), value = (string?)cell.Attribute("V");
                    if (name != null && value != null && name.StartsWith(prefix, StringComparison.Ordinal) && !cells.ContainsKey(name)) cells.Add(name, value);
                }
                reference = (string?)style.Attribute(binding);
            }
            _cache.Add(key, cells);
            return cells;
        }

        private OfficeColor Color(Dictionary<string, string> cells, string name, OfficeColor fallback) {
            if (!cells.TryGetValue(name, out string? value)) return fallback;
            if (_colors.TryGetValue(value, out string? rgb)) value = rgb;
            else if (value == "0") return OfficeColor.Black;
            else if (value == "1") return OfficeColor.White;
            return OfficeColor.TryParse(value, out OfficeColor color) ? color : fallback;
        }

        private static bool Number(Dictionary<string, string> cells, string name, out double value) {
            value = 0;
            return cells.TryGetValue(name, out string? raw) && VisioShapeGeometry.TryParseLiteralWithoutShape(raw, out value);
        }

        internal static void ApplyLocalTransparency(VisioShape shape, XElement source) {
            foreach (XElement cell in source.Elements(source.Name.Namespace + "Cell")) {
                string? name = (string?)cell.Attribute("N"), value = (string?)cell.Attribute("V");
                if (name == "LineColorTrans") shape.LineColor = Transparency(value, shape.LineColor);
                if (name == "FillForegndTrans") shape.FillColor = Transparency(value, shape.FillColor);
            }
        }

        internal static OfficeColor InheritTransparency(OfficeColor local, OfficeColor inherited) =>
            OfficeColor.FromRgba(local.R, local.G, local.B, inherited.A);

        private static OfficeColor Transparency(Dictionary<string, string> cells, string name, OfficeColor color) =>
            cells.TryGetValue(name, out string? raw) ? Transparency(raw, color) : color;

        private static OfficeColor Transparency(string? raw, OfficeColor color) =>
            VisioShapeGeometry.TryParseLiteralWithoutShape(raw, out double transparency) && transparency >= 0 && transparency <= 1
                ? OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(255 * (1 - transparency))) : color;
    }
}
