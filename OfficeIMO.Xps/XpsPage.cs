using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

/// <summary>A fixed page retaining native XML, dimensions, resource references and positioned text.</summary>
public sealed class XpsPage {
    internal XpsDocument Document { get; }
    private readonly XElement _markup;
    internal XpsPage(XpsDocument document, string partName, XElement markup) {
        Document = document; PartName = partName; _markup = markup;
        ValidatePageDimension(Width); ValidatePageDimension(Height);
    }
    /// <summary>Native package part name, without a leading slash.</summary>
    public string PartName { get; }
    /// <summary>Width in 1/96-inch units.</summary>
    public double Width => XpsPackage.Number((string?)_markup.Attribute("Width"));
    /// <summary>Height in 1/96-inch units.</summary>
    public double Height => XpsPackage.Number((string?)_markup.Attribute("Height"));
    /// <summary>A detached copy of native page markup. Use ReplaceMarkup to apply edits.</summary>
    public XElement GetMarkup() => new(_markup);
    /// <summary>Replaces page markup in the same dialect; native unsupported features remain available for preservation.</summary>
    public void ReplaceMarkup(XElement markup) {
        if (markup == null) throw new ArgumentNullException(nameof(markup));
        if (markup.Name != _markup.Name) throw new ArgumentException("Expected FixedPage in the document dialect.", nameof(markup));
        ValidatePageDimension(XpsPackage.Number((string?)markup.Attribute("Width"))); ValidatePageDimension(XpsPackage.Number((string?)markup.Attribute("Height")));
        // Reapply XML bounds before taking a caller-owned tree into the document.
        var copy = XpsPackage.Xml(XpsPackage.Serialize(markup), new XpsReadOptions(), default);
        _markup.ReplaceAttributes(copy.Attributes()); _markup.ReplaceNodes(copy.Nodes());
    }
    /// <summary>Appends native path geometry; paint strings follow the XPS color syntax.</summary>
    public XpsPage AddPath(string data, string? fill = "#FF000000", string? stroke = null, double strokeThickness = 1) {
        if (string.IsNullOrWhiteSpace(data)) throw new ArgumentException("Path geometry is required.", nameof(data));
        if (strokeThickness < 0 || double.IsNaN(strokeThickness) || double.IsInfinity(strokeThickness)) throw new ArgumentOutOfRangeException(nameof(strokeThickness));
        var element = new XElement(_markup.Name.Namespace + "Path", new XAttribute("Data", data));
        if (fill != null) element.Add(new XAttribute("Fill", fill));
        if (stroke != null) element.Add(new XAttribute("Stroke", stroke), new XAttribute("StrokeThickness", XpsPackage.N(strokeThickness)));
        _markup.Add(element); return this;
    }
    /// <summary>Appends Unicode glyph text using an embedded font. Origin is the baseline in XPS units.</summary>
    public XpsPage AddText(string text, string fontUri, double fontSize, double originX, double originY, string fill = "#FF000000") {
        if (text == null) throw new ArgumentNullException(nameof(text));
        ValidateDimension(fontSize); _ = XpsPackage.Number(XpsPackage.N(originX)); _ = XpsPackage.Number(XpsPackage.N(originY));
        _ = Document.GetPartBytes(XpsPackage.Resolve(PartName, fontUri));
        _markup.Add(new XElement(_markup.Name.Namespace + "Glyphs", new XAttribute("FontUri", fontUri), new XAttribute("FontRenderingEmSize", XpsPackage.N(fontSize)),
            new XAttribute("OriginX", XpsPackage.N(originX)), new XAttribute("OriginY", XpsPackage.N(originY)), new XAttribute("UnicodeString", text.StartsWith("{", StringComparison.Ordinal) ? "{}" + text : text), new XAttribute("Fill", fill)));
        return this;
    }
    /// <summary>Places an embedded PNG or JPEG using an image brush and a rectangular viewport.</summary>
    public XpsPage AddImage(string imageUri, double x, double y, double width, double height) {
        ValidateDimension(width); ValidateDimension(height);
        _ = XpsPackage.Number(XpsPackage.N(x)); _ = XpsPackage.Number(XpsPackage.N(y));
        string name = XpsPackage.Resolve(PartName, imageUri);
        string type = Document.ContentType(name);
        if (type != "image/png" && type != "image/jpeg") throw new NotSupportedException("The image creation API supports embedded PNG and JPEG resources.");
        var info = OfficeImageReader.Identify(Document.Part(name));
        double iw = info.Width * 96D / (info.DpiX > 0 ? info.DpiX : 96D), ih = info.Height * 96D / (info.DpiY > 0 ? info.DpiY : 96D);
        XNamespace ns = _markup.Name.Namespace;
        _markup.Add(new XElement(ns + "Path", new XAttribute("Data", "M" + XpsPackage.N(x) + "," + XpsPackage.N(y) + " h" + XpsPackage.N(width) + " v" + XpsPackage.N(height) + " h" + XpsPackage.N(-width) + " Z"),
            new XElement(ns + "Path.Fill", new XElement(ns + "ImageBrush", new XAttribute("ImageSource", imageUri),
                new XAttribute("Viewbox", "0,0," + XpsPackage.N(iw) + "," + XpsPackage.N(ih)), new XAttribute("Viewport", string.Join(",", new[] { x, y, width, height }.Select(XpsPackage.N))),
                new XAttribute("ViewboxUnits", "Absolute"), new XAttribute("ViewportUnits", "Absolute")))));
        return this;
    }
    /// <summary>Returns UnicodeString values in markup order, not inferred paragraph or reading order.</summary>
    public string ExtractText() => string.Join("\n", _markup.Descendants(_markup.Name.Namespace + "Glyphs").Select(g => Unescape((string?)g.Attribute("UnicodeString") ?? "")));
    internal static string Unescape(string text) => text.StartsWith("{}", StringComparison.Ordinal) ? text.Substring(2) : text;
    internal static void ValidatePageDimension(double value) {
        ValidateDimension(value);
        if (value < 1) throw new ArgumentOutOfRangeException(nameof(value), "XPS page dimensions must be at least one unit.");
    }
    internal static void ValidateDimension(double value) {
        if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value) || value > 100000) throw new ArgumentOutOfRangeException(nameof(value), "XPS dimensions must be positive and at most 100000 units.");
    }
    internal byte[] Serialize() => XpsPackage.Serialize(_markup);
    internal IEnumerable<string> ResourceReferences() => _markup.Descendants().Attributes().Where(a => a.Name.LocalName == "FontUri" || a.Name.LocalName == "ImageSource" || (a.Name.LocalName == "Source" && a.Parent?.Name.LocalName == "ResourceDictionary"))
        .Where(a => !a.Value.StartsWith("{", StringComparison.Ordinal)).Select(a => XpsPackage.Resolve(PartName, a.Value.Split('#')[0])).Distinct(StringComparer.OrdinalIgnoreCase);
    /// <summary>Converts the page to self-contained SVG, with explicit diagnostics for unsupported features.</summary>
    public XpsSvgResult ToSvg(bool allowPartial = false, CancellationToken cancellationToken = default) => new XpsSvgConverter(this, cancellationToken).Convert(allowPartial);
    /// <summary>Converts through the shared SVG reader. Any SVG import loss fails rather than silently disappearing.</summary>
    public OfficeDrawing ToDrawing(CancellationToken cancellationToken = default) {
        var svg = ToSvg(false, cancellationToken);
        if (!OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.Svg), new OfficeSvgDrawingReaderOptions { CancellationToken = cancellationToken, MaximumGeometryCommands = 1000000, MaximumElements = 100000 }, out var drawing, out int unsupported) || drawing == null || unsupported != 0)
            throw new NotSupportedException("The shared drawing importer cannot represent this XPS page without loss. Use ToSvg for the native vector projection.");
        return drawing;
    }
    /// <summary>Renders using the existing managed drawing codecs.</summary>
    public OfficeImageExportResult ExportImage(OfficeImageExportFormat format, OfficeImageExportOptions? options = null, CancellationToken cancellationToken = default) => ToDrawing(cancellationToken).ExportImage(format, options, cancellationToken);
}

/// <summary>A self-contained SVG projection and its explicit loss diagnostics.</summary>
public sealed class XpsSvgResult {
    internal XpsSvgResult(string svg, List<string> diagnostics) { Svg = svg; Diagnostics = diagnostics.AsReadOnly(); }
    /// <summary>Self-contained SVG. Embedded glyphs are outlined to preserve native positioning.</summary>
    public string Svg { get; }
    /// <summary>Unsupported features encountered while converting.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>Whether conversion encountered no known unsupported features.</summary>
    public bool IsComplete => Diagnostics.Count == 0;
}
