namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODT page layout and its standard master page.</summary>
public sealed class OdtPageLayout {
    private readonly OdtDocument _document;
    private readonly XElement _layout;
    private readonly XElement _properties;
    private readonly XElement _master;

    internal OdtPageLayout(OdtDocument document, XElement layout, XElement properties, XElement master) {
        _document = document;
        _layout = layout;
        _properties = properties;
        _master = master;
    }

    /// <summary>Page width.</summary>
    public OdfLength Width {
        get => ReadLength(OdfNamespaces.Fo + "page-width", "21cm", useCommonMargin: false);
        set => Set(OdfNamespaces.Fo + "page-width", value.ToString());
    }
    /// <summary>Page height.</summary>
    public OdfLength Height {
        get => ReadLength(OdfNamespaces.Fo + "page-height", "29.7cm", useCommonMargin: false);
        set => Set(OdfNamespaces.Fo + "page-height", value.ToString());
    }
    /// <summary>Top page margin.</summary>
    public OdfLength MarginTop {
        get => ReadLength(OdfNamespaces.Fo + "margin-top", "2cm");
        set => Set(OdfNamespaces.Fo + "margin-top", value.ToString());
    }
    /// <summary>Bottom page margin.</summary>
    public OdfLength MarginBottom {
        get => ReadLength(OdfNamespaces.Fo + "margin-bottom", "2cm");
        set => Set(OdfNamespaces.Fo + "margin-bottom", value.ToString());
    }
    /// <summary>Left page margin.</summary>
    public OdfLength MarginLeft {
        get => ReadLength(OdfNamespaces.Fo + "margin-left", "2cm");
        set => Set(OdfNamespaces.Fo + "margin-left", value.ToString());
    }
    /// <summary>Right page margin.</summary>
    public OdfLength MarginRight {
        get => ReadLength(OdfNamespaces.Fo + "margin-right", "2cm");
        set => Set(OdfNamespaces.Fo + "margin-right", value.ToString());
    }
    /// <summary>Master-page header content.</summary>
    public OdtHeaderFooter Header => GetHeaderFooter(OdfNamespaces.Style + "header");
    /// <summary>Master-page footer content.</summary>
    public OdtHeaderFooter Footer => GetHeaderFooter(OdfNamespaces.Style + "footer");
    /// <summary>Whether the master already defines a default header.</summary>
    public bool HasHeader => _master.Element(OdfNamespaces.Style + "header") != null;
    /// <summary>Whether the master already defines a default footer.</summary>
    public bool HasFooter => _master.Element(OdfNamespaces.Style + "footer") != null;
    /// <summary>First-page header content, when the master defines a distinct first page.</summary>
    public OdtHeaderFooter? FirstHeader => FindHeaderFooter(OdfNamespaces.Style + "header-first");
    /// <summary>First-page footer content, when the master defines a distinct first page.</summary>
    public OdtHeaderFooter? FirstFooter => FindHeaderFooter(OdfNamespaces.Style + "footer-first");
    /// <summary>Left-page header content, used for even pages in a standard left-to-right document.</summary>
    public OdtHeaderFooter? LeftHeader => FindHeaderFooter(OdfNamespaces.Style + "header-left");
    /// <summary>Left-page footer content, used for even pages in a standard left-to-right document.</summary>
    public OdtHeaderFooter? LeftFooter => FindHeaderFooter(OdfNamespaces.Style + "footer-left");

    /// <summary>Creates or returns the distinct first-page header.</summary>
    public OdtHeaderFooter EnsureFirstHeader() => GetHeaderFooter(OdfNamespaces.Style + "header-first");
    /// <summary>Creates or returns the distinct first-page footer.</summary>
    public OdtHeaderFooter EnsureFirstFooter() => GetHeaderFooter(OdfNamespaces.Style + "footer-first");
    /// <summary>Creates or returns the left-page header.</summary>
    public OdtHeaderFooter EnsureLeftHeader() => GetHeaderFooter(OdfNamespaces.Style + "header-left");
    /// <summary>Creates or returns the left-page footer.</summary>
    public OdtHeaderFooter EnsureLeftFooter() => GetHeaderFooter(OdfNamespaces.Style + "footer-left");

    internal static OdtPageLayout GetOrCreate(OdtDocument document) {
        bool changed = false;
        XDocument stylesXml = document.Package.EnsureXml("styles.xml",
            OdfPackageTemplates.CreateStyles(document.Version), "text/xml");
        XElement root = stylesXml.Root ?? throw new InvalidDataException("OpenDocument styles have no root element.");
        XElement automatic = root.Element(OdfNamespaces.Office + "automatic-styles") ?? new XElement(OdfNamespaces.Office + "automatic-styles");
        if (automatic.Parent == null) { root.Add(automatic); changed = true; }
        XElement masters = root.Element(OdfNamespaces.Office + "master-styles") ?? new XElement(OdfNamespaces.Office + "master-styles");
        if (masters.Parent == null) { root.Add(masters); changed = true; }

        XElement? master = masters.Elements(OdfNamespaces.Style + "master-page").FirstOrDefault();
        string? layoutName = (string?)master?.Attribute(OdfNamespaces.Style + "page-layout-name");
        XElement? layout = layoutName == null ? null : automatic.Elements(OdfNamespaces.Style + "page-layout")
            .FirstOrDefault(item => (string?)item.Attribute(OdfNamespaces.Style + "name") == layoutName);
        if (layout == null) {
            layoutName = "ofPage1";
            layout = new XElement(OdfNamespaces.Style + "page-layout",
                new XAttribute(OdfNamespaces.Style + "name", layoutName),
                new XElement(OdfNamespaces.Style + "page-layout-properties",
                    new XAttribute(OdfNamespaces.Fo + "page-width", "21cm"),
                    new XAttribute(OdfNamespaces.Fo + "page-height", "29.7cm"),
                    new XAttribute(OdfNamespaces.Style + "print-orientation", "portrait"),
                    new XAttribute(OdfNamespaces.Fo + "margin", "2cm")));
            automatic.Add(layout);
            changed = true;
        }
        if (master == null) {
            master = new XElement(OdfNamespaces.Style + "master-page",
                new XAttribute(OdfNamespaces.Style + "name", "Standard"),
                new XAttribute(OdfNamespaces.Style + "page-layout-name", layoutName!));
            masters.Add(master);
            changed = true;
        }
        XElement properties = layout.Element(OdfNamespaces.Style + "page-layout-properties")
            ?? new XElement(OdfNamespaces.Style + "page-layout-properties");
        if (properties.Parent == null) { layout.Add(properties); changed = true; }
        if (changed) document.MarkPartDirty("styles.xml");
        return new OdtPageLayout(document, layout, properties, master);
    }

    private OdtHeaderFooter GetHeaderFooter(XName name) {
        XElement? element = _master.Element(name);
        if (element == null) {
            if (name == OdfNamespaces.Style + "header-left" || name == OdfNamespaces.Style + "header-first")
                GetHeaderFooter(OdfNamespaces.Style + "header");
            if (name == OdfNamespaces.Style + "footer-left" || name == OdfNamespaces.Style + "footer-first")
                GetHeaderFooter(OdfNamespaces.Style + "footer");
            element = new XElement(name);
            int order = HeaderFooterOrder(name);
            XElement? following = _master.Elements().FirstOrDefault(child => HeaderFooterOrder(child.Name) > order);
            if (following == null) _master.Add(element);
            else following.AddBeforeSelf(element);
            Dirty();
        }
        return new OdtHeaderFooter(_document, element);
    }

    private static int HeaderFooterOrder(XName name) {
        if (name == OdfNamespaces.Style + "header") return 0;
        if (name == OdfNamespaces.Style + "header-left") return 1;
        if (name == OdfNamespaces.Style + "header-first") return 2;
        if (name == OdfNamespaces.Style + "footer") return 3;
        if (name == OdfNamespaces.Style + "footer-left") return 4;
        if (name == OdfNamespaces.Style + "footer-first") return 5;
        return 6;
    }

    private OdtHeaderFooter? FindHeaderFooter(XName name) {
        XElement? element = _master.Element(name);
        return element == null ? null : new OdtHeaderFooter(_document, element);
    }

    private OdfLength ReadLength(XName name, string fallback, bool useCommonMargin = true) {
        string? value = (string?)_properties.Attribute(name);
        if (value == null && useCommonMargin) value = (string?)_properties.Attribute(OdfNamespaces.Fo + "margin");
        return OdfLength.Parse(value ?? fallback);
    }

    private void Set(XName name, string value) {
        _properties.SetAttributeValue(name, value);
        Dirty();
    }

    private void Dirty() => _document.MarkPartDirty("styles.xml");
}

/// <summary>XML-backed header or footer content on an ODT master page.</summary>
public sealed class OdtHeaderFooter {
    private readonly OdtDocument _document;
    private readonly XElement _element;

    internal OdtHeaderFooter(OdtDocument document, XElement element) {
        _document = document;
        _element = element;
    }

    /// <summary>Whether this header or footer is displayed by the master page.</summary>
    public bool IsDisplayed {
        get => OdfBoolean.ReadCompatible((string?)_element.Attribute(OdfNamespaces.Style + "display"), true);
        set {
            _element.SetAttributeValue(OdfNamespaces.Style + "display", value ? null : "false");
            _document.MarkPartDirty("styles.xml");
        }
    }

    /// <summary>Paragraphs in this header or footer.</summary>
    public IReadOnlyList<OdtParagraph> Paragraphs => _element.Elements()
        .Where(element => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")
        .Select(element => new OdtParagraph(_document, element, "styles.xml")).ToList();

    /// <summary>Direct content blocks outside the paragraph and heading surface.</summary>
    public int NonParagraphBlockCount => _element.Elements()
        .Count(element => element.Name != OdfNamespaces.Text + "p" && element.Name != OdfNamespaces.Text + "h");

    /// <summary>Adds a paragraph.</summary>
    public OdtParagraph AddParagraph(string? text = null) {
        var paragraph = new XElement(OdfNamespaces.Text + "p");
        OdfTextCodec.Append(paragraph, text);
        _element.Add(paragraph);
        _document.MarkPartDirty("styles.xml");
        return new OdtParagraph(_document, paragraph, "styles.xml");
    }
}
