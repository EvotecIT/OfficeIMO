namespace OfficeIMO.OpenDocument;

/// <summary>Native OpenDocument Presentation document.</summary>
public sealed partial class OdpPresentation : OdfDocument {
    internal OdpPresentation(OdfPackage package, string? sourcePath) : base(package, sourcePath) {
        if (package.Kind != OdfDocumentKind.Presentation) throw new InvalidDataException("Package is not an OpenDocument Presentation document.");
    }

    /// <summary>Creates an empty ODF 1.4 presentation.</summary>
    public static OdpPresentation Create() => new OdpPresentation(OdfPackage.Create(OdfDocumentKind.Presentation), null);

    /// <summary>Opens an ODP document from a path.</summary>
    public static OdpPresentation Open(string path, OdfOpenOptions? options = null) {
        OdfPackage package = OdfPackage.Open(path, options, out string fullPath);
        return new OdpPresentation(package, fullPath);
    }

    /// <summary>Opens an ODP document from a stream.</summary>
    public static OdpPresentation Open(Stream stream, OdfOpenOptions? options = null) => new OdpPresentation(OdfPackage.Open(stream, options), null);

    internal XElement PresentationBody => GetBody(OdfNamespaces.Office + "presentation");

    /// <summary>Slides in presentation order.</summary>
    public IReadOnlyList<OdpSlide> Slides => PresentationBody.Elements(OdfNamespaces.Draw + "page")
        .Select(element => new OdpSlide(this, element)).ToList();

    /// <summary>Master pages stored in <c>styles.xml</c>.</summary>
    public IReadOnlyList<OdpMasterPage> MasterPages => GetStylesContainer(OdfNamespaces.Office + "master-styles")
        .Elements(OdfNamespaces.Style + "master-page").Select(element => new OdpMasterPage(this, element)).ToList();

    /// <summary>Presentation page layouts stored in <c>styles.xml</c>.</summary>
    public IReadOnlyList<OdpPresentationLayout> Layouts => GetStylesContainer(OdfNamespaces.Office + "styles")
        .Elements(OdfNamespaces.Style + "presentation-page-layout").Select(element => new OdpPresentationLayout(this, element)).ToList();

    /// <summary>Width of the default presentation page.</summary>
    public OdfLength PageWidth {
        get => GetPageLayoutProperties().ReadLength(OdfNamespaces.Fo + "page-width", "33.867cm");
        set => GetPageLayoutProperties().SetLength(OdfNamespaces.Fo + "page-width", value);
    }

    /// <summary>Height of the default presentation page.</summary>
    public OdfLength PageHeight {
        get => GetPageLayoutProperties().ReadLength(OdfNamespaces.Fo + "page-height", "19.05cm");
        set => GetPageLayoutProperties().SetLength(OdfNamespaces.Fo + "page-height", value);
    }

    /// <summary>Adds a slide using the default master and blank layout.</summary>
    public OdpSlide AddSlide(string? name = null) {
        OdpMasterPage master = EnsureDefaultMaster();
        OdpPresentationLayout layout = EnsureBlankLayout();
        string slideName = string.IsNullOrWhiteSpace(name) ? NextSlideName() : name!;
        if (Slides.Any(slide => string.Equals(slide.Name, slideName, StringComparison.Ordinal))) {
            throw new InvalidOperationException($"A slide named '{slideName}' already exists.");
        }
        var element = new XElement(OdfNamespaces.Draw + "page",
            new XAttribute(OdfNamespaces.Draw + "name", slideName),
            new XAttribute(OdfNamespaces.Draw + "master-page-name", master.Name),
            new XAttribute(OdfNamespaces.Presentation + "presentation-page-layout-name", layout.Name));
        PresentationBody.Add(element); MarkPartDirty("content.xml");
        return new OdpSlide(this, element);
    }

    /// <summary>Moves a slide to a zero-based position.</summary>
    public void MoveSlide(int sourceIndex, int destinationIndex) {
        List<XElement> slides = PresentationBody.Elements(OdfNamespaces.Draw + "page").ToList();
        if (sourceIndex < 0 || sourceIndex >= slides.Count) throw new ArgumentOutOfRangeException(nameof(sourceIndex));
        if (destinationIndex < 0 || destinationIndex >= slides.Count) throw new ArgumentOutOfRangeException(nameof(destinationIndex));
        XElement moving = slides[sourceIndex]; moving.Remove(); slides.RemoveAt(sourceIndex);
        if (destinationIndex >= slides.Count) PresentationBody.Add(moving); else slides[destinationIndex].AddBeforeSelf(moving);
        MarkPartDirty("content.xml");
    }

    /// <summary>Removes a slide by zero-based index.</summary>
    public void RemoveSlide(int index) {
        OdpSlide slide = Slides.ElementAtOrDefault(index) ?? throw new ArgumentOutOfRangeException(nameof(index));
        slide.Element.Remove(); MarkPartDirty("content.xml");
    }

    /// <summary>Adds an empty master page using the default page size.</summary>
    public OdpMasterPage AddMasterPage(string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Master page name cannot be empty.", nameof(name));
        if (MasterPages.Any(master => string.Equals(master.Name, name, StringComparison.Ordinal))) throw new InvalidOperationException($"A master page named '{name}' already exists.");
        string pageLayoutName = EnsurePageLayout().Name;
        var element = new XElement(OdfNamespaces.Style + "master-page",
            new XAttribute(OdfNamespaces.Style + "name", name),
            new XAttribute(OdfNamespaces.Style + "page-layout-name", pageLayoutName));
        GetStylesContainer(OdfNamespaces.Office + "master-styles").Add(element); MarkPartDirty("styles.xml");
        return new OdpMasterPage(this, element);
    }

    /// <summary>Adds an empty presentation layout.</summary>
    public OdpPresentationLayout AddLayout(string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Layout name cannot be empty.", nameof(name));
        if (Layouts.Any(layout => string.Equals(layout.Name, name, StringComparison.Ordinal))) throw new InvalidOperationException($"A layout named '{name}' already exists.");
        var element = new XElement(OdfNamespaces.Style + "presentation-page-layout", new XAttribute(OdfNamespaces.Style + "name", name));
        GetStylesContainer(OdfNamespaces.Office + "styles").Add(element); MarkPartDirty("styles.xml");
        return new OdpPresentationLayout(this, element);
    }

    internal XElement GetStylesContainer(XName name) {
        XElement root = GetXml("styles.xml").Root ?? throw new InvalidDataException("OpenDocument styles have no root element.");
        XElement? element = root.Element(name);
        if (element == null) { element = new XElement(name); root.Add(element); MarkPartDirty("styles.xml"); }
        return element;
    }

    private OdpPageLayout EnsurePageLayout() {
        XElement automatic = GetStylesContainer(OdfNamespaces.Office + "automatic-styles");
        XElement? element = automatic.Elements(OdfNamespaces.Style + "page-layout").FirstOrDefault();
        if (element == null) {
            element = new XElement(OdfNamespaces.Style + "page-layout",
                new XAttribute(OdfNamespaces.Style + "name", "ofSlidePage"),
                new XElement(OdfNamespaces.Style + "page-layout-properties",
                    new XAttribute(OdfNamespaces.Fo + "page-width", "33.867cm"),
                    new XAttribute(OdfNamespaces.Fo + "page-height", "19.05cm"),
                    new XAttribute(OdfNamespaces.Style + "print-orientation", "landscape"),
                    new XAttribute(OdfNamespaces.Fo + "margin", "0cm")));
            automatic.Add(element); MarkPartDirty("styles.xml");
        }
        return new OdpPageLayout(this, element);
    }

    private OdpPageLayoutProperties GetPageLayoutProperties() => EnsurePageLayout().Properties;
    private OdpMasterPage EnsureDefaultMaster() => MasterPages.FirstOrDefault() ?? AddMasterPage("Default");
    private OdpPresentationLayout EnsureBlankLayout() => Layouts.FirstOrDefault() ?? AddLayout("Blank");
    private string NextSlideName() {
        var names = new HashSet<string>(Slides.Select(slide => slide.Name), StringComparer.Ordinal);
        int index = 1; string name;
        do { name = "Slide" + index++.ToString(CultureInfo.InvariantCulture); } while (names.Contains(name));
        return name;
    }
}
