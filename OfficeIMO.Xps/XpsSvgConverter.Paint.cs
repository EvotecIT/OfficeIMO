namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private void Paint(XElement source, string property, XElement target, string attribute, Dictionary<string, Resource> scope, string part, int depth) {
        string? value = (string?)source.Attribute(property);
        XElement? brush = source.Element(source.Name.Namespace + (source.Name.LocalName + "." + property))?.Elements().SingleOrDefault();
        if (value?.StartsWith("{", StringComparison.Ordinal) == true) {
            Resource resource = ResolveResource(value, scope)!; brush = resource.Value; part = resource.Part;
        }
        if (brush == null) {
            if (value == null) { target.SetAttributeValue(attribute, "none"); return; }
            SetColor(target, attribute, attribute + "-opacity", value); return;
        }
        switch (brush.Name.LocalName) {
            case "SolidColorBrush":
                CheckAttributes(brush, "Color Opacity");
                SetColor(target, attribute, attribute + "-opacity", (string?)brush.Attribute("Color") ?? "#00000000", Unit((string?)brush.Attribute("Opacity") ?? "1"));
                break;
            case "LinearGradientBrush": case "RadialGradientBrush":
                Gradient(brush, target, attribute); break;
            case "ImageBrush": ImageBrush(brush, target, attribute, scope, part, depth); break;
            default: Loss("Brush: " + brush.Name.LocalName); target.SetAttributeValue(attribute, "none"); break;
        }
    }
    private void SetColor(XElement element, string colorAttribute, string alphaAttribute, string color, double opacity = 1) {
        if (color.StartsWith("#", StringComparison.Ordinal)) {
            string hex = color.Substring(1);
            if ((hex.Length != 3 && hex.Length != 4 && hex.Length != 6 && hex.Length != 8) || hex.Any(c => !Uri.IsHexDigit(c))) throw new InvalidDataException("Invalid XPS color.");
            if (hex.Length <= 4) hex = string.Concat(hex.Select(c => new string(c, 2)));
            if (hex.Length == 8) { opacity *= int.Parse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255D; hex = hex.Substring(2); }
            element.SetAttributeValue(colorAttribute, "#" + hex);
        } else if (color.StartsWith("sc#", StringComparison.Ordinal)) {
            var n = Numbers(color.Substring(3));
            if (n.Length != 3 && n.Length != 4) throw new InvalidDataException("Invalid scRGB color.");
            int offset = n.Length == 4 ? 1 : 0;
            if (offset == 1) opacity *= Unit(N(n[0]));
            int Channel(double c) => (int)Math.Round(Math.Max(0, Math.Min(1, c <= 0.0031308 ? c * 12.92 : 1.055 * Math.Pow(c, 1 / 2.4) - 0.055)) * 255);
            element.SetAttributeValue(colorAttribute, "rgb(" + string.Join(",", n.Skip(offset).Select(Channel)) + ")");
        } else { Loss("Color: " + (color.StartsWith("ContextColor", StringComparison.Ordinal) ? "ICC ContextColor" : "unsupported syntax")); element.SetAttributeValue(colorAttribute, "none"); }
        if (opacity != 1) element.SetAttributeValue(alphaAttribute, N(opacity));
    }
    private void Gradient(XElement brush, XElement target, string attribute) {
        CheckAttributes(brush, "StartPoint EndPoint Center GradientOrigin RadiusX RadiusY MappingMode SpreadMethod ColorInterpolationMode Opacity Transform");
        bool radial = brush.Name.LocalName == "RadialGradientBrush";
        string id = "paint" + (++_id);
        var gradient = new XElement(Svg + (radial ? "radialGradient" : "linearGradient"), new XAttribute("id", id),
            new XAttribute("gradientUnits", ((string?)brush.Attribute("MappingMode") ?? "Absolute") == "Absolute" ? "userSpaceOnUse" : "objectBoundingBox"));
        if (radial) {
            var center = Numbers((string?)brush.Attribute("Center") ?? "0,0");
            var origin = Numbers((string?)brush.Attribute("GradientOrigin") ?? "0,0");
            if (center.Length != 2 || origin.Length != 2) throw new InvalidDataException("Invalid radial gradient coordinates.");
            double rx = XpsPackage.Number((string?)brush.Attribute("RadiusX")); double ry = XpsPackage.Number((string?)brush.Attribute("RadiusY"));
            if (rx <= 0 || ry <= 0) throw new InvalidDataException("Gradient radii must be positive.");
            gradient.SetAttributeValue("cx", N(center[0])); gradient.SetAttributeValue("cy", N(center[1])); gradient.SetAttributeValue("r", N(rx));
            gradient.SetAttributeValue("fx", N(origin[0])); gradient.SetAttributeValue("fy", N(center[1] + (origin[1] - center[1]) * rx / ry));
            gradient.SetAttributeValue("gradientTransform", "translate(0 " + N(center[1]) + ") scale(1 " + N(ry / rx) + ") translate(0 " + N(-center[1]) + ")");
        } else {
            var start = Numbers((string?)brush.Attribute("StartPoint") ?? "0,0"); var end = Numbers((string?)brush.Attribute("EndPoint") ?? "1,1");
            if (start.Length != 2 || end.Length != 2) throw new InvalidDataException("Invalid gradient coordinates.");
            gradient.SetAttributeValue("x1", N(start[0])); gradient.SetAttributeValue("y1", N(start[1])); gradient.SetAttributeValue("x2", N(end[0])); gradient.SetAttributeValue("y2", N(end[1]));
        }
        string spread = (string?)brush.Attribute("SpreadMethod") ?? "Pad";
        gradient.SetAttributeValue("spreadMethod", spread.ToLowerInvariant());
        gradient.SetAttributeValue("color-interpolation", ((string?)brush.Attribute("ColorInterpolationMode") ?? "SRgbLinearInterpolation") == "ScRgbLinearInterpolation" ? "linearRGB" : "sRGB");
        if (brush.Attribute("Transform") != null || brush.Elements().Any(c => c.Name.LocalName.EndsWith(".Transform", StringComparison.Ordinal))) Loss("Gradient brush transform");
        double opacity = Unit((string?)brush.Attribute("Opacity") ?? "1");
        foreach (var child in brush.Elements()) {
            if (child.Name.LocalName != brush.Name.LocalName + ".GradientStops") { Loss("Gradient child: " + child.Name.LocalName); continue; }
            foreach (var stop in child.Elements()) {
                if (stop.Name.LocalName != "GradientStop") { Loss("Gradient stop element"); continue; }
                CheckAttributes(stop, "Offset Color");
                var svgStop = new XElement(Svg + "stop", new XAttribute("offset", N(XpsPackage.Number((string?)stop.Attribute("Offset")))));
                SetColor(svgStop, "stop-color", "stop-opacity", (string?)stop.Attribute("Color") ?? "#00000000", opacity); gradient.Add(svgStop);
            }
        }
        _defs.Add(gradient); target.SetAttributeValue(attribute, "url(#" + id + ")");
    }
    private void ImageBrush(XElement brush, XElement target, string attribute, Dictionary<string, Resource> scope, string part, int depth) {
        Charge(depth);
        CheckAttributes(brush, "ImageSource Viewbox Viewport ViewboxUnits ViewportUnits TileMode Opacity Transform");
        if (((string?)brush.Attribute("ViewboxUnits") ?? "Absolute") != "Absolute" || ((string?)brush.Attribute("ViewportUnits") ?? "Absolute") != "Absolute") { Loss("Relative image brush coordinates"); return; }
        if (((string?)brush.Attribute("TileMode") ?? "None") != "None") { Loss("Tiled image brush"); return; }
        foreach (var child in brush.Elements()) if (child.Name.LocalName != "ImageBrush.Transform") Loss("Image brush child: " + child.Name.LocalName);
        string source = (string?)brush.Attribute("ImageSource") ?? throw new InvalidDataException("Missing image source.");
        if (source.StartsWith("{", StringComparison.Ordinal)) { Loss("Color-converted image source"); return; }
        string name = XpsPackage.Resolve(part, source);
        string type = _page.Document.ContentType(name);
        if (type != "image/png" && type != "image/jpeg") { Loss("Image codec: " + type); return; }
        var viewbox = Numbers((string?)brush.Attribute("Viewbox") ?? "0,0,1,1"); var viewport = Numbers((string?)brush.Attribute("Viewport") ?? "0,0,1,1");
        if (viewbox.Length != 4 || viewport.Length != 4 || viewbox[2] <= 0 || viewbox[3] <= 0 || viewport[2] <= 0 || viewport[3] <= 0) throw new InvalidDataException("Invalid image brush rectangle.");
        byte[] bytes = _page.Document.Part(name);
        var info = OfficeIMO.Drawing.OfficeImageReader.Identify(bytes);
        double width = info.Width * 96D / (info.DpiX > 0 ? info.DpiX : 96D);
        double height = info.Height * 96D / (info.DpiY > 0 ? info.DpiY : 96D);
        string clipId = "imageClip" + (++_id);
        _defs.Add(new XElement(Svg + "clipPath", new XAttribute("id", clipId), new XElement(Svg + "rect", new XAttribute("x", N(viewport[0])), new XAttribute("y", N(viewport[1])), new XAttribute("width", N(viewport[2])), new XAttribute("height", N(viewport[3])))));
        string encoded = System.Convert.ToBase64String(bytes); _outputCharacters = checked(_outputCharacters + encoded.Length);
        if (_outputCharacters > 32 * 1024 * 1024) throw new InvalidDataException("XPS SVG output budget exceeded.");
        var image = new XElement(Svg + "image", new XAttribute("href", "data:" + type + ";base64," + encoded),
            new XAttribute("x", N(viewport[0] - viewbox[0] * viewport[2] / viewbox[2])), new XAttribute("y", N(viewport[1] - viewbox[1] * viewport[3] / viewbox[3])),
            new XAttribute("width", N(width * viewport[2] / viewbox[2])), new XAttribute("height", N(height * viewport[3] / viewbox[3])), new XAttribute("preserveAspectRatio", "none"));
        if (attribute != "fill") { Loss("Image brush stroke"); return; }
        target.SetAttributeValue("fill", "none");
        var projection = new XElement(Svg + "g", new XAttribute("clip-path", "url(#" + clipId + ")"), new XAttribute("opacity", N(Unit((string?)brush.Attribute("Opacity") ?? "1"))), image);
        string? transform = Transform(brush, scope, "Transform");
        _imageFills[target] = transform == null ? projection : new XElement(Svg + "g", new XAttribute("transform", transform), projection);
    }
}
