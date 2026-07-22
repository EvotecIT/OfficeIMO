namespace OfficeIMO.OpenDocument;

/// <summary>Indexes and creates named and automatic ODF styles without detaching them from package XML.</summary>
public sealed class OdfStyleRepository {
    private readonly OdfDocument _document;

    internal OdfStyleRepository(OdfDocument document) {
        _document = document;
    }

    /// <summary>Enumerates common styles from <c>styles.xml</c>.</summary>
    public IReadOnlyList<OdfStyle> Named => EnumerateContainer("styles.xml", OdfNamespaces.Office + "styles", false);

    /// <summary>Enumerates automatic styles from both <c>content.xml</c> and <c>styles.xml</c>.</summary>
    public IReadOnlyList<OdfStyle> Automatic {
        get {
            var styles = new List<OdfStyle>();
            styles.AddRange(EnumerateContainer("content.xml", OdfNamespaces.Office + "automatic-styles", true));
            styles.AddRange(EnumerateContainer("styles.xml", OdfNamespaces.Office + "automatic-styles", true));
            return styles;
        }
    }

    /// <summary>Finds a style by family and name, preferring content automatic styles.</summary>
    public OdfStyle? Find(OdfStyleFamily family, string name) {
        if (string.IsNullOrWhiteSpace(name)) return null;
        return Automatic.Concat(Named).FirstOrDefault(style => style.Family == family && string.Equals(style.Name, name, StringComparison.Ordinal));
    }

    /// <summary>Finds an automatic style within its owning package part before falling back to common styles.</summary>
    internal OdfStyle? FindInPart(OdfStyleFamily family, string name, string partPath) {
        if (string.IsNullOrWhiteSpace(name)) return null;
        return FindAutomaticInPart(family, name, partPath) ?? FindNamed(family, name);
    }

    /// <summary>Creates a common named style in <c>styles.xml</c>.</summary>
    public OdfStyle CreateNamed(string name, OdfStyleFamily family, string? parentStyleName = null) {
        ValidateStyleName(name);
        if (Find(family, name) != null) throw new InvalidOperationException($"A {family} style named '{name}' already exists.");
        XElement container = GetContainer("styles.xml", OdfNamespaces.Office + "styles");
        XElement element = CreateStyleElement(name, family, parentStyleName);
        container.Add(element);
        _document.MarkPartDirty("styles.xml");
        return new OdfStyle(_document, element, "styles.xml", false);
    }

    /// <summary>Creates a collision-free automatic style in <c>content.xml</c>.</summary>
    public OdfStyle CreateAutomatic(OdfStyleFamily family, string prefix = "of", string? parentStyleName = null) {
        return CreateAutomaticIn("content.xml", family, prefix, parentStyleName);
    }

    private OdfStyle CreateAutomaticIn(string partPath, OdfStyleFamily family, string prefix, string? parentStyleName) {
        if (string.IsNullOrWhiteSpace(prefix)) prefix = "of";
        string normalized = new string(prefix.Where(character => char.IsLetterOrDigit(character) || character == '_' || character == '-').ToArray());
        if (normalized.Length == 0 || !char.IsLetter(normalized[0])) normalized = "of" + normalized;
        var names = new HashSet<string>(Named.Concat(Automatic).Select(style => style.Name), StringComparer.Ordinal);
        int index = 1;
        string name;
        do {
            name = normalized + index.ToString("D4", CultureInfo.InvariantCulture);
            index++;
        } while (names.Contains(name));

        XElement container = GetContainer(partPath, OdfNamespaces.Office + "automatic-styles");
        XElement element = CreateStyleElement(name, family, parentStyleName);
        container.Add(element);
        _document.MarkPartDirty(partPath);
        return new OdfStyle(_document, element, partPath, true);
    }

    /// <summary>Resolves a style and its parent chain from most specific to least specific.</summary>
    public IReadOnlyList<OdfStyle> Resolve(OdfStyle style) {
        if (style == null) throw new ArgumentNullException(nameof(style));
        var result = new List<OdfStyle>();
        var visited = new HashSet<string>(StringComparer.Ordinal);
        OdfStyle? current = style;
        while (current != null) {
            string key = FamilyToken(current.Family) + ":" + current.Name;
            if (!visited.Add(key)) {
                _document.AddDiagnostic(new OdfDiagnostic("ODF203", OdfDiagnosticSeverity.Warning,
                    $"Style parent cycle detected at '{current.Name}'.", current.PartPath));
                break;
            }
            result.Add(current);
            current = string.IsNullOrEmpty(current.ParentStyleName) ? null : current.IsAutomatic
                ? FindInPart(current.Family, current.ParentStyleName!, current.PartPath)
                : FindNamed(current.Family, current.ParentStyleName!);
        }
        return result;
    }

    internal OdfStyle EnsureAutomaticStyle(XElement owner, XName styleAttribute, OdfStyleFamily family, string prefix, string partPath = "content.xml") {
        string? existingName = (string?)owner.Attribute(styleAttribute);
        OdfStyle? existing = existingName == null ? null : FindInPart(family, existingName, partPath);
        if (existing != null && existing.IsAutomatic && existing.PartPath == partPath &&
            IsUniquelyReferenced(owner, styleAttribute, existingName!, partPath)) return existing;

        OdfStyle created = existing != null && existing.IsAutomatic
            ? CloneAutomaticIn(partPath, family, prefix, existing)
            : CreateAutomaticIn(partPath, family, prefix, existingName);
        owner.SetAttributeValue(styleAttribute, created.Name);
        _document.MarkPartDirty(partPath);
        return created;
    }

    private OdfStyle CloneAutomaticIn(string partPath, OdfStyleFamily family, string prefix, OdfStyle source) {
        OdfStyle clone = CreateAutomaticIn(partPath, family, prefix, null);
        foreach (XAttribute attribute in source.Element.Attributes()) {
            if (attribute.Name == OdfNamespaces.Style + "name" || attribute.Name == OdfNamespaces.Style + "family") continue;
            clone.Element.SetAttributeValue(attribute.Name, attribute.Value);
        }
        clone.Element.Add(source.Element.Nodes().Select(node => CloneNode(node)));
        return clone;
    }

    private bool IsUniquelyReferenced(XElement owner, XName styleAttribute, string styleName, string partPath) {
        XDocument document = _document.GetXml(partPath);
        if (!ReferenceEquals(owner.Document, document)) return false;
        int references = document.Descendants()
            .Count(element => string.Equals((string?)element.Attribute(styleAttribute), styleName, StringComparison.Ordinal));
        return references == 1;
    }

    private OdfStyle? FindAutomaticInPart(OdfStyleFamily family, string name, string partPath) =>
        EnumerateContainer(partPath, OdfNamespaces.Office + "automatic-styles", true)
            .FirstOrDefault(style => style.Family == family && string.Equals(style.Name, name, StringComparison.Ordinal));

    private OdfStyle? FindNamed(OdfStyleFamily family, string name) => Named
        .FirstOrDefault(style => style.Family == family && string.Equals(style.Name, name, StringComparison.Ordinal));

    private static XNode CloneNode(XNode node) {
        if (node is XElement element) return new XElement(element);
        if (node is XCData cdata) return new XCData(cdata.Value);
        if (node is XText text) return new XText(text.Value);
        if (node is XComment comment) return new XComment(comment.Value);
        if (node is XProcessingInstruction instruction) return new XProcessingInstruction(instruction.Target, instruction.Data);
        throw new InvalidDataException($"Unsupported style XML node type '{node.NodeType}'.");
    }

    internal static string FamilyToken(OdfStyleFamily family) {
        switch (family) {
            case OdfStyleFamily.Text: return "text";
            case OdfStyleFamily.Paragraph: return "paragraph";
            case OdfStyleFamily.Table: return "table";
            case OdfStyleFamily.TableRow: return "table-row";
            case OdfStyleFamily.TableColumn: return "table-column";
            case OdfStyleFamily.TableCell: return "table-cell";
            case OdfStyleFamily.Graphic: return "graphic";
            case OdfStyleFamily.Presentation: return "presentation";
            case OdfStyleFamily.DrawingPage: return "drawing-page";
            case OdfStyleFamily.Chart: return "chart";
            default: throw new ArgumentOutOfRangeException(nameof(family));
        }
    }

    internal static bool TryParseFamily(string? value, out OdfStyleFamily family) {
        foreach (OdfStyleFamily candidate in global::OfficeIMO.Internal.EnumCompat.GetValues<OdfStyleFamily>()) {
            if (string.Equals(FamilyToken(candidate), value, StringComparison.Ordinal)) {
                family = candidate;
                return true;
            }
        }
        family = default;
        return false;
    }

    private IReadOnlyList<OdfStyle> EnumerateContainer(string partPath, XName containerName, bool automatic) {
        if (!_document.Package.ContainsEntry(partPath)) return Array.Empty<OdfStyle>();
        XDocument xml = _document.GetXml(partPath);
        XElement? root = xml.Root;
        XElement? container = root?.Element(containerName);
        if (container == null) return Array.Empty<OdfStyle>();
        return container.Elements(OdfNamespaces.Style + "style")
            .Where(element => TryParseFamily((string?)element.Attribute(OdfNamespaces.Style + "family"), out _))
            .Select(element => new OdfStyle(_document, element, partPath, automatic))
            .ToList();
    }

    private XElement GetContainer(string partPath, XName name) {
        XDocument xml = _document.GetXml(partPath);
        XElement root = xml.Root ?? throw new InvalidDataException($"OpenDocument part '{partPath}' has no root element.");
        XElement? container = root.Element(name);
        if (container == null) {
            container = new XElement(name);
            root.Add(container);
            _document.MarkPartDirty(partPath);
        }
        return container;
    }

    private static XElement CreateStyleElement(string name, OdfStyleFamily family, string? parentStyleName) {
        var element = new XElement(OdfNamespaces.Style + "style",
            new XAttribute(OdfNamespaces.Style + "name", name),
            new XAttribute(OdfNamespaces.Style + "family", FamilyToken(family)));
        if (!string.IsNullOrWhiteSpace(parentStyleName)) {
            element.SetAttributeValue(OdfNamespaces.Style + "parent-style-name", parentStyleName);
        }
        return element;
    }

    private static void ValidateStyleName(string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Style name cannot be empty.", nameof(name));
    }
}
