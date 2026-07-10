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

    internal static OdtPageLayout GetOrCreate(OdtDocument document) {
        bool changed = false;
        XDocument stylesXml = document.GetXml("styles.xml");
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
            element = new XElement(name);
            _master.Add(element);
            Dirty();
        }
        return new OdtHeaderFooter(_document, element);
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

    /// <summary>Paragraphs in this header or footer.</summary>
    public IReadOnlyList<OdtParagraph> Paragraphs => _element.Elements()
        .Where(element => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")
        .Select(element => new OdtParagraph(_document, element, "styles.xml")).ToList();

    /// <summary>Adds a paragraph.</summary>
    public OdtParagraph AddParagraph(string? text = null) {
        var paragraph = new XElement(OdfNamespaces.Text + "p");
        OdfTextCodec.Append(paragraph, text);
        _element.Add(paragraph);
        _document.MarkPartDirty("styles.xml");
        return new OdtParagraph(_document, paragraph, "styles.xml");
    }
}
