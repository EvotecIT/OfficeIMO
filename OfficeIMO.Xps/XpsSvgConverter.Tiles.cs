namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private readonly HashSet<XElement> _visualStack = new();
    private int _visualDepth;

    private readonly struct BrushRegion {
        internal BrushRegion(double x, double y, double width, double height) { X = x; Y = y; Width = width; Height = height; }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
    }
    private BrushRegion Rectangle(XElement brush, string name) {
        double[] values = Numbers((string?)brush.Attribute(name) ?? throw new InvalidDataException("Missing brush " + name + "."));
        if (values.Length != 4 || values[2] < 0 || values[3] < 0) throw new InvalidDataException("Invalid brush rectangle.");
        return new BrushRegion(values[0], values[1], values[2], values[3]);
    }
    private bool TileCoordinates(XElement brush, out BrushRegion source, out BrushRegion viewport) {
        if (((string?)brush.Attribute("ViewboxUnits") ?? "Absolute") != "Absolute" || ((string?)brush.Attribute("ViewportUnits") ?? "Absolute") != "Absolute")
            throw new InvalidDataException("XPS tile brush coordinates must be absolute.");
        source = Rectangle(brush, "Viewbox"); viewport = Rectangle(brush, "Viewport");
        return source.Width > 0 && source.Height > 0 && viewport.Width > 0 && viewport.Height > 0;
    }
    private void ImageBrush(XElement brush, XElement target, string attribute, Dictionary<string, Resource> scope, string part, int depth) {
        Charge(depth);
        CheckAttributes(brush, "ImageSource Viewbox Viewport ViewboxUnits ViewportUnits TileMode Opacity Transform");
        foreach (var child in brush.Elements()) if (child.Name.LocalName != "ImageBrush.Transform") Loss("Image brush child: " + child.Name.LocalName);
        if (!TileCoordinates(brush, out var source, out var viewport)) { Set(target, attribute, "none"); return; }
        string uri = (string?)brush.Attribute("ImageSource") ?? throw new InvalidDataException("Missing image source.");
        if (uri.StartsWith("{", StringComparison.Ordinal)) { Loss("Color-converted image source"); return; }
        string name = XpsPackage.Resolve(part, uri), type = _page.Document.ContentType(name);
        if (type != "image/png" && type != "image/jpeg" && type != "image/tiff") { Loss("Image codec: " + type); return; }
        byte[] bytes = _page.Document.Part(name);
        var info = OfficeIMO.Drawing.OfficeImageReader.Identify(bytes);
        double width = info.Width * 96D / (info.DpiX > 0 ? info.DpiX : 96D), height = info.Height * 96D / (info.DpiY > 0 ? info.DpiY : 96D);
        if (type == "image/tiff") {
            var metadata = OfficeIMO.Drawing.OfficeImageMetadataInspector.Inspect(bytes, info.Format, 0, _token);
            if (metadata.HasColorRenderingMetadata) { Loss("TIFF embedded color management"); return; }
            var options = new OfficeIMO.Drawing.OfficeRasterDecodeOptions { MaximumDecodedPixels = 4_000_000, CancellationToken = _token };
            if (!OfficeIMO.Drawing.OfficeRasterImageDecoder.TryDecode(bytes, options, out var raster, out _) || raster == null) {
                Loss("Unsupported TIFF encoding"); return;
            }
            bytes = OfficeIMO.Drawing.OfficeRasterImageEncoder.Encode(raster, OfficeIMO.Drawing.OfficeImageExportFormat.Png, null, 16 * 1024 * 1024, _token);
            type = "image/png";
        }
        EnsureOutputCapacity(((long)bytes.Length + 2) / 3 * 4 + 128);
        var image = Element("image", new XAttribute("href", "data:" + type + ";base64," + System.Convert.ToBase64String(bytes)),
            new XAttribute("x", "0"), new XAttribute("y", "0"), new XAttribute("width", N(width)), new XAttribute("height", N(height)), new XAttribute("preserveAspectRatio", "none"));
        ProjectTile(brush, target, attribute, image, source, viewport, scope);
    }
    private void VisualBrush(XElement brush, XElement target, string attribute, Dictionary<string, Resource> scope, string part, int depth) {
        Charge(depth);
        CheckAttributes(brush, "Visual Viewbox Viewport ViewboxUnits ViewportUnits TileMode Opacity Transform");
        foreach (var child in brush.Elements()) if (child.Name.LocalName != "VisualBrush.Visual" && child.Name.LocalName != "VisualBrush.Transform") Loss("Visual brush child: " + child.Name.LocalName);
        if (!TileCoordinates(brush, out var source, out var viewport)) { Set(target, attribute, "none"); return; }
        if (!_visualStack.Add(brush)) throw new InvalidDataException("Cyclic XPS visual brush.");
        try {
            XElement? visual = brush.Element(brush.Name.Namespace + "VisualBrush.Visual")?.Elements().SingleOrDefault();
            if (brush.Attribute("Visual") is XAttribute reference) {
                if (visual != null) throw new InvalidDataException("A visual brush cannot have both inline and referenced content.");
                Resource resource = ResolveResource(reference.Value, scope)!; visual = resource.Value; part = resource.Part;
            }
            if (visual == null) throw new InvalidDataException("Missing visual brush content.");
            if (visual.Name.NamespaceName != XpsPackage.Namespace(_page.Document.Format) || !new[] { "Canvas", "Path", "Glyphs" }.Contains(visual.Name.LocalName)) throw new InvalidDataException("Invalid visual brush content.");
            var content = Element("g");
            _visualDepth++;
            try {
                // An empty wrapper lets the ordinary native visual pipeline handle this
                // root exactly as a page child, with source-viewbox visibility for masks.
                var wrapper = new XElement(visual.Name.Namespace + "Canvas");
                RenderChildren(wrapper, content, scope, part, depth + 1, source, visual);
            } finally { _visualDepth--; }
            ProjectTile(brush, target, attribute, content, source, viewport, scope);
        } finally { _visualStack.Remove(brush); }
    }
    private void ProjectTile(XElement brush, XElement target, string attribute, XElement content, BrushRegion source, BrushRegion viewport, Dictionary<string, Resource> scope) {
        double sx = viewport.Width / source.Width, sy = viewport.Height / source.Height;
        string mapping = "matrix(" + N(sx) + " 0 0 " + N(sy) + " " + N(viewport.X - source.X * sx) + " " + N(viewport.Y - source.Y * sy) + ")";
        var mapped = Element("g", new XAttribute("transform", mapping), content);
        string clipId = "tileClip" + (++_id);
        _defs.Add(Element("clipPath", new XAttribute("id", clipId), Element("rect", new XAttribute("x", N(viewport.X)), new XAttribute("y", N(viewport.Y)), new XAttribute("width", N(viewport.Width)), new XAttribute("height", N(viewport.Height)))));
        var tile = Element("g", new XAttribute("clip-path", "url(#" + clipId + ")"), new XAttribute("opacity", N(Unit((string?)brush.Attribute("Opacity") ?? "1"))), mapped);
        string mode = (string?)brush.Attribute("TileMode") ?? "None";
        string? transform = Transform(brush, scope, "Transform");
        if (mode == "None") {
            XElement projected = transform == null ? tile : Element("g", new XAttribute("transform", transform), tile);
            if (attribute == "stroke") {
                // Keep a visible placeholder until Stroke has validated and attached
                // native width/dash/cap properties. ApplyBrushFill turns it into coverage.
                Set(target, "stroke", "#ffffff"); _brushStrokes[target] = projected;
            } else {
                Set(target, "fill", "none"); _brushFills[target] = projected;
            }
            return;
        }
        if (!new[] { "Tile", "FlipX", "FlipY", "FlipXY" }.Contains(mode)) throw new InvalidDataException("Invalid tile mode.");
        bool flipX = mode == "FlipX" || mode == "FlipXY", flipY = mode == "FlipY" || mode == "FlipXY";
        string id = "tile" + (++_id);
        var pattern = Element("pattern", new XAttribute("id", id), new XAttribute("patternUnits", "userSpaceOnUse"), new XAttribute("patternContentUnits", "userSpaceOnUse"),
            new XAttribute("x", N(viewport.X)), new XAttribute("y", N(viewport.Y)), new XAttribute("width", N(viewport.Width * (flipX ? 2 : 1))), new XAttribute("height", N(viewport.Height * (flipY ? 2 : 1))), tile);
        string mirrorX = "translate(" + N(2 * (viewport.X + viewport.Width)) + " 0) scale(-1 1)";
        string mirrorY = "translate(0 " + N(2 * (viewport.Y + viewport.Height)) + ") scale(1 -1)";
        if (flipX) pattern.Add(Element("g", new XAttribute("transform", mirrorX), CloneProjection(tile)));
        if (flipY) pattern.Add(Element("g", new XAttribute("transform", mirrorY), CloneProjection(tile)));
        if (flipX && flipY) pattern.Add(Element("g", new XAttribute("transform", mirrorX + " " + mirrorY), CloneProjection(tile)));
        if (transform != null) Set(pattern, "patternTransform", transform);
        _defs.Add(pattern); Set(target, attribute, "url(#" + id + ")");
    }
    private XElement CloneProjection(XElement source) {
        var clone = Element(source.Name.LocalName, source.Attributes().Select(a => (object)new XAttribute(a)).ToArray());
        foreach (var child in source.Elements()) clone.Add(CloneProjection(child));
        return clone;
    }
}
