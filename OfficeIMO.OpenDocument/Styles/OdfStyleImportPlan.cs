namespace OfficeIMO.OpenDocument;

/// <summary>Copies a reachable ODF style graph while retaining named/automatic part scope and source defaults.</summary>
internal sealed partial class OdfStyleImportPlan {
    private readonly OdfDocument _source;
    private readonly OdfDocument _destination;
    private readonly OdfResourceImportPlan _resources;
    private readonly Dictionary<XElement, string> _mapped = new Dictionary<XElement, string>();
    private readonly List<(string Part, XName Container, XElement Element)> _copies = new List<(string, XName, XElement)>();
    private readonly Dictionary<(string Part, string Family), string> _defaultNames = new Dictionary<(string, string), string>();
    private readonly HashSet<string> _names;
    private readonly HashSet<string> _checkedDefaults = new HashSet<string>(StringComparer.Ordinal);
    private readonly HashSet<XElement> _validatedParents = new HashSet<XElement>();
    private string? _sourceMaster, _destinationMaster;
    private int _next = 1;
    private int _importDepth;
    internal OdfStyleImportPlan(OdfDocument source, OdfDocument destination, OdfResourceImportPlan resources) {
        _source = source; _destination = destination; _resources = resources;
        _names = new HashSet<string>(new[] { "content.xml", "styles.xml" }.SelectMany(part => destination.GetXml(part).Descendants())
            .SelectMany(element => element.Attributes().Where(attribute => attribute.Name == OdfNamespaces.Style + "name" || attribute.Name == OdfNamespaces.Draw + "name"))
            .Select(attribute => attribute.Value), StringComparer.Ordinal);
    }
    internal string ReserveName() {
        string name;
        do { name = "odgImportStyle" + _next++.ToString(CultureInfo.InvariantCulture); } while (!_names.Add(name));
        return name;
    }
    internal void SetMasterMapping(string source, string destination) { _sourceMaster = source; _destinationMaster = destination; }

    internal void Apply() {
        foreach (var copy in _copies) {
            XElement root = _destination.GetXml(copy.Part).Root!;
            XElement container = OdfXmlContainers.Ensure(root, copy.Container);
            container.Add(copy.Element); _destination.MarkPartDirty(copy.Part);
        }
    }

    private XElement Container(string part, XName name) => _source.GetXml(part).Root?.Element(name) ?? new XElement(name);
    private static XElement? Unique(IEnumerable<XElement> candidates, string name, XName attribute, string kind) {
        XElement[] matches = candidates.Where(element => (string?)element.Attribute(attribute) == name).Take(2).ToArray();
        if (matches.Length > 1) throw new InvalidDataException("Imported " + kind + " is ambiguous: " + name);
        return matches.FirstOrDefault();
    }
    private string ImportStyle(string name, string family, string part, bool commonOnly = false) {
        XElement? style = commonOnly ? null : Unique(Container(part, OdfNamespaces.Office + "automatic-styles")
            .Elements(OdfNamespaces.Style + "style").Where(element => (string?)element.Attribute(OdfNamespaces.Style + "family") == family),
            name, OdfNamespaces.Style + "name", family + " style");
        string ownerPart = part;
        XName ownerContainer = OdfNamespaces.Office + "automatic-styles";
        if (style == null) {
            ownerPart = "styles.xml"; ownerContainer = OdfNamespaces.Office + "styles";
            style = Unique(Container(ownerPart, ownerContainer).Elements(OdfNamespaces.Style + "style")
                .Where(element => (string?)element.Attribute(OdfNamespaces.Style + "family") == family), name, OdfNamespaces.Style + "name", family + " style");
        }
        if (style == null) throw new InvalidDataException("Imported " + family + " style is missing: " + name);
        CheckDefaults(family);
        ValidateParentChain(style);
        return ImportDefinition(style, ownerPart, ownerContainer);
    }

    private void ValidateParentChain(XElement style) {
        var visited = new HashSet<XElement>();
        XElement? current = style;
        while (current != null) {
            if (_validatedParents.Contains(current)) break;
            if (!visited.Add(current)) throw new InvalidDataException("Imported style parent chain contains a cycle.");
            string? parent = (string?)current.Attribute(OdfNamespaces.Style + "parent-style-name");
            if (string.IsNullOrEmpty(parent)) break;
            string family = (string?)current.Attribute(OdfNamespaces.Style + "family") ?? "";
            XElement? next = Unique(Container("styles.xml", OdfNamespaces.Office + "styles").Elements(OdfNamespaces.Style + "style")
                .Where(element => (string?)element.Attribute(OdfNamespaces.Style + "family") == family), parent!, OdfNamespaces.Style + "name", "parent style");
            current = next ?? throw new InvalidDataException("Imported style parent is missing: " + parent);
        }
        _validatedParents.UnionWith(visited);
    }

    private string ImportDefinition(XElement source, string part, XName container) {
        if (_mapped.TryGetValue(source, out string? name)) return name;
        if (_importDepth >= 256) throw new NotSupportedException("Imported style dependencies exceed the supported depth of 256.");
        name = ReserveName(); _mapped.Add(source, name);
        XElement copy = new XElement(source);
        CaptureLeaderTextOrigins(copy, part, container == OdfNamespaces.Office + "styles");
        XName nameAttribute = source.Attribute(OdfNamespaces.Draw + "name") != null ? OdfNamespaces.Draw + "name" : OdfNamespaces.Style + "name";
        copy.SetAttributeValue(nameAttribute, name);
        if (source.Name == OdfNamespaces.Style + "style" && source.Attribute(OdfNamespaces.Style + "parent-style-name") == null &&
            (string?)source.Attribute(OdfNamespaces.Style + "family") is not ("paragraph" or "text"))
            MergeDefault(copy, (string?)source.Attribute(OdfNamespaces.Style + "family") ?? "");
        if (_source.Version != _destination.Version && (_source.Version == OdfVersion.V1_2 || _destination.Version == OdfVersion.V1_2) &&
            (source.Name == OdfNamespaces.Draw + "gradient" || source.Name == OdfNamespaces.Draw + "opacity") &&
            (string?)source.Attribute(OdfNamespaces.Draw + "style") != "radial" &&
            double.TryParse((string?)source.Attribute(OdfNamespaces.Draw + "angle"), NumberStyles.Float, CultureInfo.InvariantCulture, out double angle) && angle != 0)
            throw new NotSupportedException("Drawing import cannot reinterpret a nonzero unitless gradient angle across ODF 1.2. Use explicit angle units.");
        _copies.Add((part, container, copy));
        _importDepth++;
        try { Rewrite(copy, part); _resources.Rewrite(copy); }
        finally { _importDepth--; }
        return name;
    }

    private XElement? Default(OdfDocument document, string family) => Unique(document.GetXml("styles.xml").Root?
        .Element(OdfNamespaces.Office + "styles")?.Elements(OdfNamespaces.Style + "default-style")
        .Where(element => (string?)element.Attribute(OdfNamespaces.Style + "family") == family) ?? Enumerable.Empty<XElement>(),
        family, OdfNamespaces.Style + "family", "default style");

    private void CheckDefaults(string family) {
        if (!_checkedDefaults.Add(family)) return;
        XElement? source = Default(_source, family), destination = Default(_destination, family);
        // Copying explicit source defaults can override different destination values, but an unspecified
        // source value cannot cancel a destination-only default without a property-specific reset rule.
        foreach (XElement properties in destination?.Elements() ?? Enumerable.Empty<XElement>()) {
            XElement? sourceProperties = source?.Element(properties.Name);
            if (properties.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && sourceProperties?.Attribute(attribute.Name) == null) ||
                properties.Elements().Any(element => sourceProperties?.Element(element.Name) == null))
                throw new NotSupportedException("Drawing import cannot isolate destination-only " + family + " defaults. Use a destination with compatible defaults.");
        }
    }

    private void MergeDefault(XElement style, string family) {
        XElement? defaults = Default(_source, family);
        foreach (XElement properties in defaults?.Elements() ?? Enumerable.Empty<XElement>()) {
            XElement? local = style.Element(properties.Name);
            if (local == null) {
                local = new XElement(properties); style.Add(local);
                CaptureLeaderTextOrigins(local, "styles.xml", commonOnly: true);
                foreach (XAttribute font in local.DescendantsAndSelf().Attributes().Where(attribute => IsFontReference(attribute.Name))) _fontOrigins[font] = "styles.xml";
                continue;
            }
            var explicitGroups = new HashSet<string>(local.Attributes().Select(attribute => PropertyGroup(properties.Name, attribute.Name)), StringComparer.Ordinal);
            CaptureEdgeFallbacks(local, properties, new[] { local });
            foreach (XAttribute attribute in properties.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration))
                if (!IsEdgeProperty(attribute.Name) && !explicitGroups.Contains(PropertyGroup(properties.Name, attribute.Name))) {
                    var copy = new XAttribute(attribute); local.Add(copy);
                    if (IsFontReference(copy.Name)) _fontOrigins[copy] = "styles.xml";
                }
            foreach (XElement child in properties.Elements()) if (local.Element(child.Name) == null) {
                var copy = new XElement(child); local.Add(copy);
                CaptureLeaderTextOrigins(copy, "styles.xml", commonOnly: true);
            }
        }
    }

    private string ImportDefault(string family, string part) {
        CheckDefaults(family);
        if (_defaultNames.TryGetValue((part, family), out string? name)) return name;
        name = ReserveName(); _defaultNames.Add((part, family), name);
        var copy = new XElement(OdfNamespaces.Style + "style", new XAttribute(OdfNamespaces.Style + "name", name),
            new XAttribute(OdfNamespaces.Style + "family", family));
        MergeDefault(copy, family);
        _copies.Add((part, OdfNamespaces.Office + "automatic-styles", copy));
        Rewrite(copy, part); _resources.Rewrite(copy);
        return name;
    }

    private string ImportNamedDefinition(string name, string kind, string part, bool commonOnly = false) {
        if (kind == "data") {
            OdfDataStyle? data = commonOnly ? _source.Styles.FindCommonDataStyle(name) : _source.Styles.FindDataStyle(name, part);
            return data != null ? ImportDefinition(data.Element, data.PartPath, data.IsAutomatic ? OdfNamespaces.Office + "automatic-styles" : OdfNamespaces.Office + "styles")
                : throw new InvalidDataException("Imported data definition is missing: " + name);
        }
        XElement? definition = null;
        string ownerPart = part; XName ownerContainer = OdfNamespaces.Office + "automatic-styles";
        foreach (var location in new[] { (part, OdfNamespaces.Office + "automatic-styles"), ("styles.xml", OdfNamespaces.Office + "styles") }) {
            definition = Unique(Container(location.Item1, location.Item2).Elements().Where(element => MatchesKind(element, kind)), name,
                kind == "list" || kind == "data" ? OdfNamespaces.Style + "name" : OdfNamespaces.Draw + "name", kind + " definition");
            if (definition != null) { ownerPart = location.Item1; ownerContainer = location.Item2; break; }
        }
        return definition != null ? ImportDefinition(definition, ownerPart, ownerContainer)
            : throw new InvalidDataException("Imported " + kind + " definition is missing: " + name);
    }
    private static bool MatchesKind(XElement element, string kind) => kind switch {
        "list" => element.Name == OdfNamespaces.Text + "list-style",
        "data" => element.Name.Namespace == OdfNamespaces.Number && element.Attribute(OdfNamespaces.Style + "name") != null,
        "gradient" => element.Name == OdfNamespaces.Draw + "gradient" || element.Name == OdfNamespaces.Svg + "linearGradient" || element.Name == OdfNamespaces.Svg + "radialGradient",
        _ => element.Name == OdfNamespaces.Draw + kind
    };
    private string ImportFont(string name, string part) {
        foreach (string candidate in new[] { part, "styles.xml", "content.xml" }.Distinct(StringComparer.Ordinal)) {
            XElement? face = Unique(Container(candidate, OdfNamespaces.Office + "font-face-decls").Elements(OdfNamespaces.Style + "font-face"),
                name, OdfNamespaces.Style + "name", "font face");
            if (face != null) return ImportDefinition(face, candidate, OdfNamespaces.Office + "font-face-decls");
        }
        throw new InvalidDataException("Imported font face is missing: " + name);
    }
}
