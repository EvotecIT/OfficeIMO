namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    private XNamespace ResourceKeyNamespace => _page.Document.Format == XpsFormat.Xps ? "http://schemas.microsoft.com/winfx/2006/xaml" : "http://schemas.openxps.org/oxps/v1.0/resourcedictionary-key";
    private readonly XpsPage _page;
    private readonly CancellationToken _token;
    private readonly List<string> _diagnostics = new();
    private readonly XElement _defs;
    private readonly Dictionary<XElement, XElement> _imageFills = new();
    private int _id;
    private int _visited;
    private int _points;
    private int _pathCommands;
    private long _outputCharacters;
    private int _outputNodes;
    private readonly Dictionary<string, XElement> _resourceDictionaries = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, OfficeIMO.Drawing.OfficeTrueTypeFont> _fonts = new(StringComparer.OrdinalIgnoreCase);
    internal XpsSvgConverter(XpsPage page, CancellationToken token) { _page = page; _token = token; _defs = Element("defs"); }
    private sealed class Resource {
        internal Resource(XElement value, string part) { Value = value; Part = part; }
        internal XElement Value { get; }
        internal string Part { get; }
    }
    internal XpsSvgResult Convert(bool allowPartial) {
        _token.ThrowIfCancellationRequested();
        var root = Element("svg", new XAttribute("width", N(_page.Width)), new XAttribute("height", N(_page.Height)), new XAttribute("viewBox", "0 0 " + N(_page.Width) + " " + N(_page.Height)), _defs);
        XElement page = _page.GetMarkup();
        CheckAttributes(page, "Width Height ContentBox BleedBox Name");
        if (page.Attribute("Name") is XAttribute pageName) Set(root, "id", "xps-" + pageName.Value);
        RenderChildren(page, root, new Dictionary<string, Resource>(), _page.PartName, 0);
        if (_diagnostics.Count != 0 && !allowPartial) throw new NotSupportedException("XPS conversion would lose features: " + string.Join("; ", _diagnostics.Take(12)));
        return new XpsSvgResult(root.ToString(SaveOptions.DisableFormatting), _diagnostics);
    }
    private static string N(double value) => XpsPackage.N(value);
    private void Loss(string message) { if (_diagnostics.Count < 100 && !_diagnostics.Contains(message)) _diagnostics.Add(message); }
    private void CheckAttributes(XElement e, string allowed) {
        var names = new HashSet<string>((allowed + " Name").Split(' '), StringComparer.Ordinal);
        foreach (var a in e.Attributes()) {
            if (a.IsNamespaceDeclaration || a.Name == XNamespace.Xml + "lang" || a.Name == ResourceKeyNamespace + "Key") continue;
            if ((a.Name.NamespaceName.Length != 0 && a.Name.Namespace != e.Name.Namespace) || !names.Contains(a.Name.LocalName)) Loss(e.Name.LocalName + "." + a.Name.LocalName);
        }
    }
    private void Charge(int depth) {
        _token.ThrowIfCancellationRequested();
        if (depth > 64 || ++_visited > 100000) throw new InvalidDataException("XPS conversion complexity limit exceeded.");
    }
    private Dictionary<string, Resource> Resources(XElement parent, Dictionary<string, Resource> inherited, string part, int depth) {
        XElement? container = parent.Element(parent.Name.Namespace + (parent.Name.LocalName + ".Resources"));
        if (container == null) return inherited;
        ChargeBindings(inherited.Count);
        var scope = new Dictionary<string, Resource>(inherited, StringComparer.Ordinal);
        var localKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var dictionary in container.Elements()) ReadDictionary(dictionary, part, scope, localKeys, new HashSet<string>(StringComparer.OrdinalIgnoreCase), depth);
        return scope;
    }
    private void ReadDictionary(XElement dictionary, string part, Dictionary<string, Resource> scope, HashSet<string> keys, HashSet<string> stack, int depth) {
        Charge(depth);
        if (dictionary.Name != XName.Get("ResourceDictionary", XpsPackage.Namespace(_page.Document.Format))) { Loss("Unknown resource dictionary"); return; }
        CheckAttributes(dictionary, "Source");
        string? source = (string?)dictionary.Attribute("Source");
        if (source != null) {
            string name = XpsPackage.Resolve(part, source);
            if (!stack.Add(name)) throw new InvalidDataException("Cyclic XPS resource dictionary.");
            if (_page.Document.ContentType(name) != XpsPackage.Type("resourcedictionary")) throw new InvalidDataException("Invalid resource dictionary content type.");
            if (!_resourceDictionaries.TryGetValue(name, out var external)) {
                external = _page.Document.ReadXml(name, _token);
                _resourceDictionaries.Add(name, external);
            }
            ReadDictionary(external, name, scope, keys, stack, depth + 1);
            stack.Remove(name);
        }
        foreach (var item in dictionary.Elements()) {
            string? key = (string?)item.Attribute(ResourceKeyNamespace + "Key");
            if (key == null) { Loss("Unkeyed resource: " + item.Name.LocalName); continue; }
            if (!keys.Add(key)) throw new InvalidDataException("Duplicate XPS resource key.");
            ChargeBindings(1);
            scope[key] = new Resource(item, part);
        }
    }
    private Resource? ResolveResource(string value, Dictionary<string, Resource> resources) {
        const string prefix = "{StaticResource ";
        if (!value.StartsWith(prefix, StringComparison.Ordinal) || !value.EndsWith("}", StringComparison.Ordinal)) throw new InvalidDataException("Unsupported XPS markup extension.");
        string key = value.Substring(prefix.Length, value.Length - prefix.Length - 1).Trim();
        if (!resources.TryGetValue(key, out var found)) throw new InvalidDataException("Missing XPS resource: " + key);
        return found;
    }
    private void RenderChildren(XElement parent, XElement target, Dictionary<string, Resource> inherited, string part, int depth) {
        Charge(depth);
        var scope = Resources(parent, inherited, part, depth);
        foreach (XElement child in parent.Elements()) {
            _token.ThrowIfCancellationRequested();
            if (child.Name.LocalName == parent.Name.LocalName + ".Resources") continue;
            if (child.Name.LocalName == parent.Name.LocalName + ".RenderTransform" || child.Name.LocalName == parent.Name.LocalName + ".Clip") continue;
            if (child.Name.NamespaceName != XpsPackage.Namespace(_page.Document.Format)) { Loss("Foreign element: " + child.Name); continue; }
            XElement? result;
            switch (child.Name.LocalName) {
                case "Canvas":
                    CheckAttributes(child, "RenderTransform Clip Opacity FixedPage.NavigateUri");
                    result = Element("g");
                    RenderChildren(child, result, scope, part, depth + 1); break;
                case "Path": result = PathElement(child, scope, part, depth + 1); break;
                case "Glyphs": result = Glyphs(child, scope, part, depth + 1); break;
                default: Loss("Element: " + child.Name.LocalName); continue;
            }
            if (result == null) continue;
            string? transform = Transform(child, scope);
            if (transform != null) Set(result, "transform", transform);
            string? clip = Geometry(child, "Clip", scope);
            if (clip != null) {
                string id = "clip" + (++_id);
                string path = StripFillRule(clip, out string rule);
                _defs.Add(Element("clipPath", new XAttribute("id", id), new XAttribute("clipPathUnits", "userSpaceOnUse"), Element("path", new XAttribute("d", path), new XAttribute("clip-rule", rule))));
                Set(result, "clip-path", "url(#" + id + ")");
            }
            if (child.Attribute("Opacity") is XAttribute opacity) Set(result, "opacity", N(Unit(opacity.Value)));
            string? link = (string?)child.Attribute("FixedPage.NavigateUri");
            if (link != null) {
                string? href = Link(link, part);
                if (href != null) result = Element("a", new XAttribute("href", href), result);
            }
            if (child.Attribute("Name") is XAttribute name) Set(result, "id", "xps-" + name.Value);
            target.Add(result);
        }
    }
    private string? Link(string target, string part) {
        if (target.StartsWith("#", StringComparison.Ordinal)) return "#xps-" + target.Substring(1);
        if (Uri.TryCreate(target, UriKind.Absolute, out var absolute) && !target.StartsWith("/", StringComparison.Ordinal)) {
            if (absolute.Scheme == "https" || absolute.Scheme == "http" || absolute.Scheme == "mailto") return target;
            Loss("Unsafe navigation URI"); return null;
        }
        string[] pieces = target.Split('#');
        string name = XpsPackage.Resolve(part, pieces[0]);
        string? anchor = pieces.Length == 2 ? pieces[1] : null;
        int index = _page.Document.LinkTargetPage(name, anchor);
        if (index < 0) { Loss("Navigation to unresolved page target"); return null; }
        // Unknown advertised names and numeric page addresses refer to the top of the page.
        string fragment = anchor != null && _page.Document.Pages[index].HasNamedTarget(anchor) ? "#xps-" + anchor : "";
        return string.Equals(_page.Document.Pages[index].PartName, _page.PartName, StringComparison.OrdinalIgnoreCase) ? fragment : "page-" + (index + 1).ToString(CultureInfo.InvariantCulture) + ".svg" + fragment;
    }
    private static double Unit(string value) {
        double number = XpsPackage.Number(value);
        if (number < 0 || number > 1) throw new InvalidDataException("XPS opacity must be between 0 and 1.");
        return number;
    }
    private string? Transform(XElement element, Dictionary<string, Resource> resources, string property = "RenderTransform") {
        string? value = (string?)element.Attribute(property);
        XElement? matrix = element.Element(element.Name.Namespace + (element.Name.LocalName + "." + property))?.Elements().SingleOrDefault();
        if (value?.StartsWith("{", StringComparison.Ordinal) == true) matrix = ResolveResource(value, resources)!.Value;
        if (matrix != null) {
            if (matrix.Name.LocalName != "MatrixTransform") { Loss("Non-matrix transform"); return null; }
            CheckAttributes(matrix, "Matrix"); value = (string?)matrix.Attribute("Matrix");
        }
        if (value == null) return null;
        var numbers = Numbers(value);
        if (numbers.Length != 6) throw new InvalidDataException("An XPS matrix requires six coefficients.");
        return "matrix(" + string.Join(" ", numbers.Select(N)) + ")";
    }
    private static double[] Numbers(string value) => value.Split(new[] { ',', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(v => XpsPackage.Number(v)).ToArray();
    private string StripFillRule(string path, out string rule) {
        path = path.Trim(); rule = "evenodd";
        if (path.StartsWith("F", StringComparison.Ordinal)) {
            int index = 1;
            while (index < path.Length && char.IsWhiteSpace(path[index])) index++;
            if (index >= path.Length || (path[index] != '0' && path[index] != '1')) throw new InvalidDataException("Invalid XPS fill rule.");
            rule = path[index] == '1' ? "nonzero" : "evenodd";
            path = path.Substring(index + 1).Trim();
        }
        if (path.Length > 0) {
            if (!OfficeIMO.Drawing.OfficeSvgPathDataParser.TryParse(path, 100000 - _pathCommands, out var commands, out _, allowEmptyGeometry: true))
                throw new InvalidDataException("Malformed or excessive XPS path geometry.");
            _pathCommands += commands.Count;
        }
        return path;
    }
    private string? Geometry(XElement e, string property, Dictionary<string, Resource> scope) {
        string? value = (string?)e.Attribute(property);
        XElement? geometry = e.Element(e.Name.Namespace + (e.Name.LocalName + "." + property))?.Elements().SingleOrDefault();
        if (value?.StartsWith("{", StringComparison.Ordinal) == true) geometry = ResolveResource(value, scope)!.Value;
        if (geometry == null) return value;
        if (geometry.Name.LocalName != "PathGeometry") { Loss("Geometry: " + geometry.Name.LocalName); return null; }
        CheckAttributes(geometry, "Figures FillRule");
        string figures = (string?)geometry.Attribute("Figures") ?? Figures(geometry);
        return ((string?)geometry.Attribute("FillRule") == "NonZero" ? "F1 " : "F0 ") + figures;
    }
    private string Figures(XElement geometry) {
        var data = new StringBuilder();
        foreach (var figure in geometry.Elements()) {
            if (figure.Name.LocalName != "PathFigure") { Loss("Geometry child: " + figure.Name.LocalName); continue; }
            CheckAttributes(figure, "StartPoint IsClosed IsFilled");
            if ((string?)figure.Attribute("IsFilled") == "false") Loss("Unfilled path figure");
            data.Append("M ").Append((string?)figure.Attribute("StartPoint") ?? throw new InvalidDataException("Missing figure start."));
            foreach (var segment in figure.Elements()) {
                CheckAttributes(segment, "Point Points Point1 Point2 Point3 Size RotationAngle IsLargeArc SweepDirection IsStroked");
                if ((string?)segment.Attribute("IsStroked") == "false") Loss("Unstroked path segment");
                switch (segment.Name.LocalName) {
                    case "PolyLineSegment": data.Append(" L ").Append((string?)segment.Attribute("Points")); break;
                    case "PolyBezierSegment": data.Append(" C ").Append((string?)segment.Attribute("Points")); break;
                    case "PolyQuadraticBezierSegment": data.Append(" Q ").Append((string?)segment.Attribute("Points")); break;
                    case "ArcSegment": data.Append(" A ").Append((string?)segment.Attribute("Size")).Append(' ').Append((string?)segment.Attribute("RotationAngle") ?? "0").Append(' ').Append((string?)segment.Attribute("IsLargeArc") == "true" ? "1" : "0").Append(' ').Append((string?)segment.Attribute("SweepDirection") == "Clockwise" ? "1" : "0").Append(' ').Append((string?)segment.Attribute("Point")); break;
                    default: Loss("Path segment: " + segment.Name.LocalName); break;
                }
            }
            if ((string?)figure.Attribute("IsClosed") == "true") data.Append(" Z ");
        }
        return data.ToString();
    }
    private XElement PathElement(XElement e, Dictionary<string, Resource> scope, string part, int depth) {
        Charge(depth);
        CheckAttributes(e, "Data Fill Stroke StrokeThickness StrokeDashArray StrokeDashOffset StrokeStartLineCap StrokeEndLineCap StrokeDashCap StrokeLineJoin StrokeMiterLimit RenderTransform Clip Opacity FixedPage.NavigateUri");
        string path = StripFillRule(Geometry(e, "Data", scope) ?? "", out string rule);
        var result = Element("path", new XAttribute("d", path), new XAttribute("fill-rule", rule));
        Paint(e, "Fill", result, "fill", scope, part, depth);
        Paint(e, "Stroke", result, "stroke", scope, part, depth);
        Stroke(e, result, path);
        foreach (var child in e.Elements()) if (!new[] { "Path.Data", "Path.Fill", "Path.Stroke", "Path.Clip", "Path.RenderTransform" }.Contains(child.Name.LocalName)) Loss(child.Name.LocalName);
        return ApplyImageFill(result);
    }
    private XElement ApplyImageFill(XElement path) {
        if (!_imageFills.TryGetValue(path, out var image)) return path;
        string id = "pathClip" + (++_id);
        _defs.Add(Element("clipPath", new XAttribute("id", id), Element("path", new XAttribute("d", (string?)path.Attribute("d") ?? ""), new XAttribute("clip-rule", (string?)path.Attribute("fill-rule") ?? "evenodd"))));
        return Element("g", Element("g", new XAttribute("clip-path", "url(#" + id + ")"), image), path);
    }
}
