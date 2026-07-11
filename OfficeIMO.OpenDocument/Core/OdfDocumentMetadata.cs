namespace OfficeIMO.OpenDocument;

/// <summary>OpenDocument metadata backed by `meta.xml`.</summary>
public sealed class OdfDocumentMetadata {
    private readonly OdfDocument _owner;
    private readonly XElement _metadata;

    internal OdfDocumentMetadata(OdfDocument owner) {
        _owner = owner;
        XDocument document = owner.Package.EnsureXml("meta.xml", OdfPackageTemplates.CreateMetadata(owner.Version), "text/xml");
        XElement root = document.Root ?? throw new InvalidDataException("OpenDocument metadata has no root element.");
        _metadata = root.Element(OdfNamespaces.Office + "meta") ?? new XElement(OdfNamespaces.Office + "meta");
        if (_metadata.Parent == null) {
            root.Add(_metadata);
            owner.MarkPartDirty("meta.xml");
        }
    }

    /// <summary>Document title.</summary>
    public string? Title { get => Get(OdfNamespaces.Dc + "title"); set => Set(OdfNamespaces.Dc + "title", value); }
    /// <summary>Document subject.</summary>
    public string? Subject { get => Get(OdfNamespaces.Dc + "subject"); set => Set(OdfNamespaces.Dc + "subject", value); }
    /// <summary>Document description.</summary>
    public string? Description { get => Get(OdfNamespaces.Dc + "description"); set => Set(OdfNamespaces.Dc + "description", value); }
    /// <summary>Initial creator.</summary>
    public string? Creator { get => Get(OdfNamespaces.Meta + "initial-creator"); set => Set(OdfNamespaces.Meta + "initial-creator", value); }
    /// <summary>Primary document language.</summary>
    public string? Language { get => Get(OdfNamespaces.Dc + "language"); set => Set(OdfNamespaces.Dc + "language", value); }
    /// <summary>Producer recorded in the package.</summary>
    public string? Generator { get => Get(OdfNamespaces.Meta + "generator"); set => Set(OdfNamespaces.Meta + "generator", value); }

    /// <summary>Creation timestamp.</summary>
    public DateTimeOffset? CreationDate {
        get => GetDate(OdfNamespaces.Meta + "creation-date");
        set => SetDate(OdfNamespaces.Meta + "creation-date", value);
    }

    /// <summary>Last modification timestamp.</summary>
    public DateTimeOffset? ModifiedDate {
        get => GetDate(OdfNamespaces.Dc + "date");
        set => SetDate(OdfNamespaces.Dc + "date", value);
    }

    /// <summary>Gets a custom `meta:user-defined` value.</summary>
    public string? GetCustomProperty(string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Custom property name cannot be empty.", nameof(name));
        XElement? element = _metadata.Elements(OdfNamespaces.Meta + "user-defined")
            .FirstOrDefault(item => string.Equals((string?)item.Attribute(OdfNamespaces.Meta + "name"), name, StringComparison.Ordinal));
        return element?.Value;
    }

    /// <summary>Adds, replaces, or removes a custom metadata value.</summary>
    public void SetCustomProperty(string name, string? value) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Custom property name cannot be empty.", nameof(name));
        List<XElement> matches = _metadata.Elements(OdfNamespaces.Meta + "user-defined")
            .Where(item => string.Equals((string?)item.Attribute(OdfNamespaces.Meta + "name"), name, StringComparison.Ordinal))
            .ToList();
        if (value == null) {
            foreach (XElement match in matches) match.Remove();
        } else if (matches.Count > 0) {
            matches[0].Value = value;
            foreach (XElement duplicate in matches.Skip(1)) duplicate.Remove();
        } else {
            _metadata.Add(new XElement(OdfNamespaces.Meta + "user-defined",
                new XAttribute(OdfNamespaces.Meta + "name", name), value));
        }
        _owner.MarkPartDirty("meta.xml");
    }

    private string? Get(XName name) => _metadata.Element(name)?.Value;

    private void Set(XName name, string? value) {
        XElement? element = _metadata.Element(name);
        if (value == null) {
            element?.Remove();
        } else if (element == null) {
            _metadata.Add(new XElement(name, value));
        } else {
            element.Value = value;
        }
        _owner.MarkPartDirty("meta.xml");
    }

    private DateTimeOffset? GetDate(XName name) {
        string? value = Get(name);
        return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTimeOffset parsed) ? parsed : (DateTimeOffset?)null;
    }

    private void SetDate(XName name, DateTimeOffset? value) => Set(name, value?.ToString("o", CultureInfo.InvariantCulture));
}
