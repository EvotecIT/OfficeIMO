namespace OfficeIMO.OpenDocument;

/// <summary>Native OpenDocument Presentation document.</summary>
public sealed partial class OdpPresentation : OdfDocument {
    internal OdpPresentation(OdfPackage package, string? sourcePath) : base(package, sourcePath) {
        if (package.Kind != OdfDocumentKind.Presentation) throw new InvalidDataException("Package is not an OpenDocument Presentation document.");
    }

    /// <summary>Creates an empty ODF 1.4 presentation.</summary>
    public static OdpPresentation Create() => new OdpPresentation(OdfPackage.Create(OdfDocumentKind.Presentation), null);

    /// <summary>Loads an ODP document from a path.</summary>
    public new static OdpPresentation Load(string path, OdfLoadOptions? options = null) {
        OdfPackage package = OdfPackage.Load(path, options, out string fullPath);
        return new OdpPresentation(package, fullPath);
    }

    /// <summary>Loads an ODP document from a stream.</summary>
    public new static OdpPresentation Load(Stream stream, OdfLoadOptions? options = null) => new OdpPresentation(OdfPackage.Load(stream, options), null);

    /// <summary>Asynchronously loads an ODP document from a path.</summary>
    public new static async Task<OdpPresentation> LoadAsync(string path, OdfLoadOptions? options = null, CancellationToken cancellationToken = default) {
        OdfDocument document = await OdfDocument.LoadAsync(path, options, cancellationToken).ConfigureAwait(false);
        return document as OdpPresentation ?? throw new InvalidDataException("Package is not an OpenDocument Presentation document.");
    }

    /// <summary>Asynchronously loads an ODP document from a caller-owned stream.</summary>
    public new static async Task<OdpPresentation> LoadAsync(Stream stream, OdfLoadOptions? options = null, CancellationToken cancellationToken = default) {
        OdfPackage package = await LoadPackageAsync(stream, options, cancellationToken).ConfigureAwait(false);
        return new OdpPresentation(package, null);
    }

    internal XElement PresentationBody => GetBody(OdfNamespaces.Office + "presentation");

    /// <summary>Slides in presentation order.</summary>
    public IReadOnlyList<OdpSlide> Slides => PresentationBody.Elements(OdfNamespaces.Draw + "page")
        .Select(element => new OdpSlide(this, element)).ToList();

    /// <summary>Master pages stored in <c>styles.xml</c>.</summary>
    public IReadOnlyList<OdpMasterPage> MasterPages => GetExistingXml("styles.xml")?.Root?
        .Element(OdfNamespaces.Office + "master-styles")?.Elements(OdfNamespaces.Style + "master-page")
        .Select(element => new OdpMasterPage(this, element)).ToArray() ?? Array.Empty<OdpMasterPage>();

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
        string slideName;
        if (string.IsNullOrWhiteSpace(name)) {
            slideName = NextSlideName();
        } else {
            slideName = name!;
            if (FindSlideElement(slideName) != null) {
                throw new InvalidOperationException($"A slide named '{slideName}' already exists.");
            }
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
        ArgumentOutOfRangeException.ThrowIfNegative(sourceIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(sourceIndex, slides.Count);
        ArgumentOutOfRangeException.ThrowIfNegative(destinationIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(destinationIndex, slides.Count);
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
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
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
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        if (Layouts.Any(layout => string.Equals(layout.Name, name, StringComparison.Ordinal))) throw new InvalidOperationException($"A layout named '{name}' already exists.");
        var element = new XElement(OdfNamespaces.Style + "presentation-page-layout", new XAttribute(OdfNamespaces.Style + "name", name));
        GetStylesContainer(OdfNamespaces.Office + "styles").Add(element); MarkPartDirty("styles.xml");
        return new OdpPresentationLayout(this, element);
    }

    internal XElement GetStylesContainer(XName name) {
        XElement root = Package.EnsureXml("styles.xml", OdfPackageTemplates.CreateStyles(Version), "text/xml").Root
            ?? throw new InvalidDataException("OpenDocument styles have no root element.");
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

    private OdpPageLayoutProperties GetPageLayoutProperties() {
        OdpMasterPage? master = null;
        string? referencedMaster = Slides.Select(slide => slide.MasterPageName)
            .FirstOrDefault(name => !string.IsNullOrWhiteSpace(name));
        if (referencedMaster != null) {
            master = MasterPages.FirstOrDefault(item => string.Equals(item.Name, referencedMaster, StringComparison.Ordinal));
        }
        master ??= MasterPages.FirstOrDefault();
        string? pageLayoutName = master?.PageLayoutName;
        if (!string.IsNullOrWhiteSpace(pageLayoutName)) {
            XElement? referenced = GetStylesContainer(OdfNamespaces.Office + "automatic-styles")
                .Elements(OdfNamespaces.Style + "page-layout")
                .FirstOrDefault(element => string.Equals((string?)element.Attribute(OdfNamespaces.Style + "name"),
                    pageLayoutName, StringComparison.Ordinal));
            if (referenced != null) return new OdpPageLayout(this, referenced).Properties;
        }
        return EnsurePageLayout().Properties;
    }
    private OdpMasterPage EnsureDefaultMaster() => MasterPages.FirstOrDefault() ?? AddMasterPage("Default");
    private OdpPresentationLayout EnsureBlankLayout() => Layouts.FirstOrDefault() ?? AddLayout("Blank");
    private XElement? FindSlideElement(string slideName) {
        foreach (XElement page in PresentationBody.Elements(OdfNamespaces.Draw + "page")) {
            if (string.Equals((string?)page.Attribute(OdfNamespaces.Draw + "name"), slideName, StringComparison.Ordinal)) return page;
        }
        return null;
    }

    /// <summary>Returns the lowest unused "SlideN" name in one pass over the slide elements.</summary>
    private string NextSlideName() {
        // With n slides, one of Slide1..Slide(n+1) is always free, so only those indexes are tracked.
        int slideCount = PresentationBody.Elements(OdfNamespaces.Draw + "page").Count();
        var used = new System.Collections.BitArray(slideCount + 2);
        foreach (XElement page in PresentationBody.Elements(OdfNamespaces.Draw + "page")) {
            int index = ParseGeneratedSlideIndex((string?)page.Attribute(OdfNamespaces.Draw + "name"));
            if (index > 0 && index < used.Length) used[index] = true;
        }
        int next = 1;
        while (used[next]) next++;
        return "Slide" + next.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>Returns N for a canonical "SlideN" name (no sign or leading zero), otherwise 0.</summary>
    private static int ParseGeneratedSlideIndex(string? name) {
        const string prefix = "Slide";
        if (name == null || name.Length <= prefix.Length || name.Length > prefix.Length + 9 ||
            !name.StartsWith(prefix, StringComparison.Ordinal) || name[prefix.Length] == '0') return 0;
        int value = 0;
        for (int i = prefix.Length; i < name.Length; i++) {
            char c = name[i];
            if (c < '0' || c > '9') return 0;
            value = value * 10 + (c - '0');
        }
        return value;
    }
}
