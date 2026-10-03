namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    /// <summary>Primary title. Other titles and refinements are retained.</summary>
    public string Title { get => GetDc("title"); set => SetDc("title", value); }
    /// <summary>Primary language. Chapter-level language declarations are edited independently.</summary>
    public string Language { get => GetDc("language"); set => SetDc("language", value); }
    /// <summary>Selected package identity. Identity changes are rejected for obfuscated/encrypted packages.</summary>
    public string Identifier {
        get => IdentifierElement?.Value ?? string.Empty;
        set {
            RequireText(value, nameof(value));
            if (_encryption.Count != 0 && value != _originalIdentifier) throw new NotSupportedException("Changing encrypted/obfuscated package identity requires re-keying resources.");
            XElement element = IdentifierElement ?? throw new InvalidDataException("Selected package identifier is missing.");
            EditPackageElement(element, proposed => proposed.Value = value);
        }
    }
    /// <summary>Primary creator, preserving other contributors and refinements.</summary>
    public string? Creator {
        get => RequireSection("metadata").Element(Dc + "creator")?.Value;
        set {
            if (value == null) throw new ArgumentNullException(nameof(value));
            SetDc("creator", value);
        }
    }
    /// <summary>Page progression declaration, such as ltr or rtl; null restores the package default.</summary>
    public string? PageProgressionDirection {
        get => (string?)RequireSection("spine").Attribute("page-progression-direction");
        set {
            if (value != null && value != "ltr" && value != "rtl" && value != "default") throw new ArgumentOutOfRangeException(nameof(value));
            if (PackageVersion == "2.0" && value != null) throw new NotSupportedException("Page progression is an EPUB 3 declaration.");
            EditPackageElement(RequireSection("spine"), proposed => proposed.SetAttributeValue("page-progression-direction", value));
        }
    }
    /// <summary>Adds an ordered Dublin Core value; existing unknown metadata remains intact.</summary>
    public void AddDublinCoreMetadata(string name, string value, string? id = null, string? language = null) {
        XmlConvert.VerifyNCName(name); RequireText(value, nameof(value));
        if (name == "language") EpubLanguageTag.Require(value, nameof(value));
        if (language != null) EpubLanguageTag.Require(language, nameof(language));
        if (id != null) VerifyAvailableId(id);
        var element = new XElement(Dc + name, value);
        element.SetAttributeValue("id", id); element.SetAttributeValue(XNamespace.Xml + "lang", language);
        EditPackageElement(RequireSection("metadata"), proposed => proposed.Add(new XElement(element)));
    }
    /// <summary>Sets one EPUB 3 property/refinement without replacing unrelated declarations.</summary>
    public void SetMetadataProperty(string property, string value, string? refines = null) {
        ValidateMetadataProperty(property, value, refines);
        XElement metadata = RequireSection("metadata");
        XElement? existing = metadata.Elements(Opf + "meta").FirstOrDefault(element =>
            EpubVocabulary.Expand(Root, (string?)element.Attribute("property") ?? string.Empty) == EpubVocabulary.Expand(Root, property) &&
            (string?)element.Attribute("refines") == refines);
        if (existing != null) EditPackageElement(existing, proposed => proposed.Value = value);
        else AddMetadataProperty(property, value, refines);
    }
    /// <summary>Adds an ordered EPUB 3 property value, including repeatable accessibility declarations.</summary>
    public void AddMetadataProperty(string property, string value, string? refines = null) {
        ValidateMetadataProperty(property, value, refines);
        var element = new XElement(Opf + "meta", new XAttribute("property", property),
            refines == null ? null : new XAttribute("refines", refines), value);
        EditPackageElement(RequireSection("metadata"), proposed => proposed.Add(new XElement(element)));
    }
    private void ValidateMetadataProperty(string property, string value, string? refines) {
        if (PackageVersion != "3.0") throw new NotSupportedException("EPUB 3 property metadata is unavailable in OPF 2.");
        RequireText(property, nameof(property)); RequireText(value, nameof(value));
        EpubVocabulary.ValidatePropertyName(Root, property);
        if (refines != null && (!refines.StartsWith("#", StringComparison.Ordinal) ||
            !Root.DescendantsAndSelf().Any(element => (string?)element.Attribute("id") == refines.Substring(1))))
            throw new ArgumentException("A refinement must target an existing package id.", nameof(refines));
    }
    /// <summary>Declares a custom vocabulary prefix without replacing other declarations.</summary>
    public void DeclareVocabularyPrefix(string prefix, string vocabularyUri) {
        if (PackageVersion != "3.0") throw new NotSupportedException("Vocabulary declarations require EPUB 3.");
        EpubVocabulary.ValidateDeclaration(prefix, vocabularyUri);
        string current = (string?)Root.Attribute("prefix") ?? string.Empty;
        string[] parts = current.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Where((_, index) => index % 2 == 0).Contains(prefix + ":")) throw new ArgumentException("Vocabulary prefix already declared.", nameof(prefix));
        EditPackageElement(Root, proposed => proposed.SetAttributeValue("prefix", (current + " " + prefix + ": " + vocabularyUri).Trim()));
    }
    /// <summary>Sets typed package layout metadata. This declares layout; it does not create fixed-page geometry.</summary>
    public void SetRenditionLayout(EpubRenditionLayout layout) {
        if (!Enum.IsDefined(typeof(EpubRenditionLayout), layout)) throw new ArgumentOutOfRangeException(nameof(layout));
        if (EpubVocabulary.Expand(Root, "rendition:layout") != "http://www.idpf.org/vocab/rendition/#layout")
            throw new InvalidOperationException("The rendition prefix has been reassigned to a different vocabulary.");
        SetMetadataProperty("rendition:layout", layout == EpubRenditionLayout.PrePaginated ? "pre-paginated" : "reflowable");
    }

    private XElement? IdentifierElement => RequireSection("metadata").Elements(Dc + "identifier")
        .FirstOrDefault(element => (string?)element.Attribute("id") == (string?)Root.Attribute("unique-identifier"));
    private string GetDc(string name) => RequireSection("metadata").Element(Dc + name)?.Value ?? string.Empty;
    private void SetDc(string name, string value) {
        RequireText(value, nameof(value));
        if (name == "language") EpubLanguageTag.Require(value, nameof(value));
        XElement? element = RequireSection("metadata").Element(Dc + name);
        if (element == null) EditPackageElement(RequireSection("metadata"), proposed => proposed.Add(new XElement(Dc + name, value)));
        else EditPackageElement(element, proposed => proposed.Value = value);
    }
    private void VerifyAvailableId(string id) {
        XmlConvert.VerifyNCName(id);
        if (Root.DescendantsAndSelf().Any(element => (string?)element.Attribute("id") == id)) throw new ArgumentException("Package id already exists: " + id);
    }
}
