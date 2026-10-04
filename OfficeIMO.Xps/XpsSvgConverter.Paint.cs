namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private void Paint(XElement source, string property, XElement target, string attribute, Dictionary<string, Resource> scope, string part, int depth) {
        string? value = (string?)source.Attribute(property);
        XElement? brush = source.Element(source.Name.Namespace + (source.Name.LocalName + "." + property))?.Elements().SingleOrDefault();
        if (value?.StartsWith("{", StringComparison.Ordinal) == true) {
            Resource resource = ResolveResource(value, scope)!; brush = resource.Value; part = resource.Part;
        }
        if (brush == null) {
            if (value == null) { Set(target, attribute, "none"); return; }
            SetColor(target, attribute, attribute + "-opacity", value); return;
        }
        switch (brush.Name.LocalName) {
            case "SolidColorBrush":
                CheckAttributes(brush, "Color Opacity");
                SetColor(target, attribute, attribute + "-opacity", (string?)brush.Attribute("Color") ?? "#00000000", Unit((string?)brush.Attribute("Opacity") ?? "1"));
                break;
            case "LinearGradientBrush": case "RadialGradientBrush":
                Gradient(brush, target, attribute, scope); break;
            case "VisualBrush": VisualBrush(brush, target, attribute, scope, part, depth); break;
            case "ImageBrush": ImageBrush(brush, target, attribute, scope, part, depth); break;
            default: Loss("Brush: " + brush.Name.LocalName); Set(target, attribute, "none"); break;
        }
    }
    private void SetColor(XElement element, string colorAttribute, string alphaAttribute, string color, double opacity = 1) {
        if (color.StartsWith("#", StringComparison.Ordinal)) {
            string hex = color.Substring(1);
            if ((hex.Length != 3 && hex.Length != 4 && hex.Length != 6 && hex.Length != 8) || hex.Any(c => !Uri.IsHexDigit(c))) throw new InvalidDataException("Invalid XPS color.");
            if (hex.Length <= 4) hex = string.Concat(hex.Select(c => new string(c, 2)));
            if (hex.Length == 8) { opacity *= int.Parse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255D; hex = hex.Substring(2); }
            Set(element, colorAttribute, "#" + hex);
        } else if (color.StartsWith("sc#", StringComparison.Ordinal)) {
            var n = Numbers(color.Substring(3));
            if (n.Length != 3 && n.Length != 4) throw new InvalidDataException("Invalid scRGB color.");
            int offset = n.Length == 4 ? 1 : 0;
            if (offset == 1) opacity *= Unit(N(n[0]));
            int Channel(double c) => (int)Math.Round(Math.Max(0, Math.Min(1, c <= 0.0031308 ? c * 12.92 : 1.055 * Math.Pow(c, 1 / 2.4) - 0.055)) * 255);
            Set(element, colorAttribute, "rgb(" + string.Join(",", n.Skip(offset).Select(Channel)) + ")");
        } else { Loss("Color: " + (color.StartsWith("ContextColor", StringComparison.Ordinal) ? "ICC ContextColor" : "unsupported syntax")); Set(element, colorAttribute, "none"); }
        if (opacity != 1) Set(element, alphaAttribute, N(opacity));
    }
    private void Gradient(XElement brush, XElement target, string attribute, Dictionary<string, Resource> scope) {
        CheckAttributes(brush, "StartPoint EndPoint Center GradientOrigin RadiusX RadiusY MappingMode SpreadMethod ColorInterpolationMode Opacity Transform");
        bool radial = brush.Name.LocalName == "RadialGradientBrush";
        string id = "paint" + (++_id);
        var gradient = Element((radial ? "radialGradient" : "linearGradient"), new XAttribute("id", id),
            new XAttribute("gradientUnits", ((string?)brush.Attribute("MappingMode") ?? "Absolute") == "Absolute" ? "userSpaceOnUse" : "objectBoundingBox"));
        if (radial) {
            var center = Numbers((string?)brush.Attribute("Center") ?? "0,0");
            var origin = Numbers((string?)brush.Attribute("GradientOrigin") ?? "0,0");
            if (center.Length != 2 || origin.Length != 2) throw new InvalidDataException("Invalid radial gradient coordinates.");
            double rx = XpsPackage.Number((string?)brush.Attribute("RadiusX")); double ry = XpsPackage.Number((string?)brush.Attribute("RadiusY"));
            if (rx <= 0 || ry <= 0) throw new InvalidDataException("Gradient radii must be positive.");
            Set(gradient, "cx", N(center[0])); Set(gradient, "cy", N(center[1])); Set(gradient, "r", N(rx));
            Set(gradient, "fx", N(origin[0])); Set(gradient, "fy", N(center[1] + (origin[1] - center[1]) * rx / ry));
            Set(gradient, "gradientTransform", "translate(0 " + N(center[1]) + ") scale(1 " + N(ry / rx) + ") translate(0 " + N(-center[1]) + ")");
        } else {
            var start = Numbers((string?)brush.Attribute("StartPoint") ?? "0,0"); var end = Numbers((string?)brush.Attribute("EndPoint") ?? "1,1");
            if (start.Length != 2 || end.Length != 2) throw new InvalidDataException("Invalid gradient coordinates.");
            Set(gradient, "x1", N(start[0])); Set(gradient, "y1", N(start[1])); Set(gradient, "x2", N(end[0])); Set(gradient, "y2", N(end[1]));
        }
        string spread = (string?)brush.Attribute("SpreadMethod") ?? "Pad";
        Set(gradient, "spreadMethod", spread.ToLowerInvariant());
        Set(gradient, "color-interpolation", ((string?)brush.Attribute("ColorInterpolationMode") ?? "SRgbLinearInterpolation") == "ScRgbLinearInterpolation" ? "linearRGB" : "sRGB");
        string? transform = Transform(brush, scope, "Transform");
        if (transform != null) Set(gradient, "gradientTransform", transform + " " + ((string?)gradient.Attribute("gradientTransform") ?? ""));
        double opacity = Unit((string?)brush.Attribute("Opacity") ?? "1");
        foreach (var child in brush.Elements()) {
            if (child.Name.LocalName == brush.Name.LocalName + ".Transform") continue;
            if (child.Name.LocalName != brush.Name.LocalName + ".GradientStops") { Loss("Gradient child: " + child.Name.LocalName); continue; }
            foreach (var stop in child.Elements()) {
                if (stop.Name.LocalName != "GradientStop") { Loss("Gradient stop element"); continue; }
                CheckAttributes(stop, "Offset Color");
                var svgStop = Element("stop", new XAttribute("offset", N(XpsPackage.Number((string?)stop.Attribute("Offset")))));
                SetColor(svgStop, "stop-color", "stop-opacity", (string?)stop.Attribute("Color") ?? "#00000000", opacity); gradient.Add(svgStop);
            }
        }
        _defs.Add(gradient); Set(target, attribute, "url(#" + id + ")");
    }
}
