namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed OpenDocument drawing with editable pages and preserved package content.</summary>
public sealed partial class OdgDocument : OdfDocument {
    internal OdgDocument(OdfPackage package, string? sourcePath) : base(package, sourcePath) {
        if (package.Kind != OdfDocumentKind.Graphics) throw new InvalidDataException("Document is not an OpenDocument drawing.");
    }

    /// <summary>Creates an empty ODF 1.4 drawing.</summary>
    public static OdgDocument Create() => new OdgDocument(OdfPackage.Create(OdfDocumentKind.Graphics), null);
    /// <summary>Loads an ODG package from a path.</summary>
    public new static OdgDocument Load(string path, OdfLoadOptions? options = null) {
        OdfPackage package = OdfPackage.Load(path, options, out string fullPath);
        return new OdgDocument(package, fullPath);
    }
    /// <summary>Loads an ODG package from a caller-owned stream.</summary>
    public new static OdgDocument Load(Stream stream, OdfLoadOptions? options = null) => new OdgDocument(OdfPackage.Load(stream, options), null);
    /// <summary>Loads a flat OpenDocument drawing (FODG).</summary>
    public new static OdgDocument LoadFlatXml(string path, OdfLoadOptions? options = null) =>
        OdfDocument.LoadFlatXml(path, options) as OdgDocument ?? throw new InvalidDataException("Flat document is not an OpenDocument drawing.");
    /// <summary>Loads FODG from a caller-owned stream.</summary>
    public new static OdgDocument LoadFlatXml(Stream stream, OdfLoadOptions? options = null) =>
        OdfDocument.LoadFlatXml(stream, options) as OdgDocument ?? throw new InvalidDataException("Flat document is not an OpenDocument drawing.");
    /// <summary>Asynchronously loads an ODG package from a path.</summary>
    public new static async Task<OdgDocument> LoadAsync(string path, OdfLoadOptions? options = null, CancellationToken cancellationToken = default) =>
        await OdfDocument.LoadAsync(path, options, cancellationToken).ConfigureAwait(false) as OdgDocument
            ?? throw new InvalidDataException("Document is not an OpenDocument drawing.");
    /// <summary>Asynchronously loads an ODG package from a caller-owned stream.</summary>
    public new static async Task<OdgDocument> LoadAsync(Stream stream, OdfLoadOptions? options = null, CancellationToken cancellationToken = default) =>
        new OdgDocument(await LoadPackageAsync(stream, options, cancellationToken).ConfigureAwait(false), null);

    /// <summary>Document-wide layers, inherited by pages without their own or master layer set.</summary>
    public OdgLayers Layers => new OdgLayers(this, EnsureContainer(GetXml("styles.xml").Root!, OdfNamespaces.Office + "master-styles"), "styles.xml");

    internal XElement DrawingBody => GetBody(OdfNamespaces.Office + "drawing");
    /// <summary>Drawing pages in document order.</summary>
    public IReadOnlyList<OdgPage> Pages => DrawingBody.Elements(OdfNamespaces.Draw + "page").Select(element => new OdgPage(this, element)).ToList();

    /// <summary>Adds a page with its own layout and master. Dimensions default to A4 portrait.</summary>
    public OdgPage AddPage(string? name = null, OdfLength? width = null, OdfLength? height = null) {
        string pageName = name ?? NextName(Pages.Select(page => page.Name), "Page");
        if (string.IsNullOrWhiteSpace(pageName)) throw new ArgumentException("Page name cannot be empty.", nameof(name));
        if (Pages.Any(page => page.Name == pageName)) throw new ArgumentException("Page names must be unique.", nameof(name));
        OdfLength w = width ?? OdfLength.Centimeters(21), h = height ?? OdfLength.Centimeters(29.7);
        ValidateDimension(w); ValidateDimension(h);
        XElement styles = Package.EnsureXml("styles.xml", OdfPackageTemplates.CreateStyles(Version), "text/xml").Root!;
        XElement automatic = EnsureContainer(styles, OdfNamespaces.Office + "automatic-styles");
        XElement masters = EnsureContainer(styles, OdfNamespaces.Office + "master-styles");
        string layout = NextName(automatic.Elements().Select(e => (string?)e.Attribute(OdfNamespaces.Style + "name") ?? ""), "ofDrawLayout");
        string master = NextName(masters.Elements().Select(e => (string?)e.Attribute(OdfNamespaces.Style + "name") ?? ""), "ofDrawMaster");
        automatic.AddFirst(new XElement(OdfNamespaces.Style + "page-layout", new XAttribute(OdfNamespaces.Style + "name", layout),
            new XElement(OdfNamespaces.Style + "page-layout-properties", new XAttribute(OdfNamespaces.Fo + "page-width", w),
                new XAttribute(OdfNamespaces.Fo + "page-height", h), new XAttribute(OdfNamespaces.Fo + "margin", "0cm"),
                new XAttribute(OdfNamespaces.Style + "print-orientation", w.ToPoints() > h.ToPoints() ? "landscape" : "portrait"))));
        masters.Add(new XElement(OdfNamespaces.Style + "master-page", new XAttribute(OdfNamespaces.Style + "name", master),
            new XAttribute(OdfNamespaces.Style + "page-layout-name", layout)));
        var element = new XElement(OdfNamespaces.Draw + "page", new XAttribute(OdfNamespaces.Draw + "name", pageName),
            new XAttribute(OdfNamespaces.Draw + "master-page-name", master));
        DrawingBody.Add(element); MarkPartDirty("content.xml"); MarkPartDirty("styles.xml");
        return new OdgPage(this, element);
    }

    /// <summary>Removes a page at a zero-based position.</summary>
    public void RemovePage(int index) {
        OdgPage page = Pages.ElementAtOrDefault(index) ?? throw new ArgumentOutOfRangeException(nameof(index));
        page.Element.Remove(); MarkPartDirty("content.xml");
    }
    /// <summary>Moves a page to a zero-based position.</summary>
    public void MovePage(int sourceIndex, int destinationIndex) {
        var pages = DrawingBody.Elements(OdfNamespaces.Draw + "page").ToList();
        if (sourceIndex < 0 || sourceIndex >= pages.Count) throw new ArgumentOutOfRangeException(nameof(sourceIndex));
        if (destinationIndex < 0 || destinationIndex >= pages.Count) throw new ArgumentOutOfRangeException(nameof(destinationIndex));
        XElement page = pages[sourceIndex]; page.Remove(); pages.RemoveAt(sourceIndex);
        if (destinationIndex == pages.Count) DrawingBody.Add(page); else pages[destinationIndex].AddBeforeSelf(page);
        MarkPartDirty("content.xml");
    }
    internal static string NextName(IEnumerable<string> existing, string prefix) {
        var names = new HashSet<string>(existing, StringComparer.Ordinal);
        int index = 1;
        while (names.Contains(prefix + index.ToString(CultureInfo.InvariantCulture))) index++;
        return prefix + index.ToString(CultureInfo.InvariantCulture);
    }
    internal static void ValidateDimension(OdfLength value) {
        if (!value.TryToPoints(out double points) || points <= 0) throw new ArgumentOutOfRangeException(nameof(value), "Page dimensions must be positive absolute lengths.");
    }
    private static XElement EnsureContainer(XElement root, XName name) => OdfXmlContainers.Ensure(root, name);
}
