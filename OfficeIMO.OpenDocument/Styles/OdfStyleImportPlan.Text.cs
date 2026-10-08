namespace OfficeIMO.OpenDocument;

internal sealed partial class OdfStyleImportPlan {
    private readonly Dictionary<XElement, XElement> _textSources = new Dictionary<XElement, XElement>();
    private readonly Dictionary<XAttribute, string> _fontOrigins = new Dictionary<XAttribute, string>();
    private readonly Dictionary<string, string> _textSnapshots = new Dictionary<string, string>(StringComparer.Ordinal);

    /// <summary>Captures text fallbacks in their original graphic context rather than promoting family defaults.</summary>
    internal void RewriteContent(XElement source, XElement copy, string part) {
        XElement[] originals = source.DescendantsAndSelf().ToArray(), copies = copy.DescendantsAndSelf().ToArray();
        if (originals.Length != copies.Length) throw new InvalidDataException("Drawing copy changed before style resolution.");
        if (originals.SelectMany(element => element.Attributes()).Any(attribute => attribute.Name.LocalName == "class-names" &&
            (attribute.Name.Namespace == OdfNamespaces.Draw || attribute.Name.Namespace == OdfNamespaces.Text || attribute.Name.Namespace == OdfNamespaces.Presentation)))
            throw new NotSupportedException("Drawing import cannot isolate multiple class-style bindings.");
        for (int i = 0; i < originals.Length; i++) if (IsTextSnapshotOwner(originals[i])) _textSources.Add(copies[i], originals[i]);
        Rewrite(copy, part);
    }
    private static bool IsTextSnapshotOwner(XElement element) => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h" ||
        element.Name == OdfNamespaces.Text + "span" || element.Name == OdfNamespaces.Text + "a";
    private static bool IsFontReference(XName name) => name.Namespace == OdfNamespaces.Style && name.LocalName is "font-name" or "font-name-asian" or "font-name-complex";
    private static string TextPropertyGroup(XName name) {
        if (name == OdfNamespaces.Fo + "font-family" || name == OdfNamespaces.Style + "font-name") return "font-family";
        foreach (string script in new[] { "asian", "complex" })
            if (name == OdfNamespaces.Style + ("font-family-" + script) || name == OdfNamespaces.Style + ("font-name-" + script)) return "font-family-" + script;
        foreach (string decoration in new[] { "underline", "line-through" })
            if (name == OdfNamespaces.Style + ("text-" + decoration + "-type") || name == OdfNamespaces.Style + ("text-" + decoration + "-style")) return "decoration-" + decoration;
        return name.ToString();
    }
    private static string PropertyGroup(XName properties, XName name) =>
        properties == OdfNamespaces.Style + "text-properties" ? TextPropertyGroup(name) :
        properties == OdfNamespaces.Style + "graphic-properties" &&
            (name == OdfNamespaces.Draw + "fit-to-size" || name == OdfNamespaces.Style + "shrink-to-fit") ? "text-fitting" : name.ToString();
    private static bool HasTextProperty(XElement properties, XName name) => properties.Attributes().Any(attribute => PropertyGroup(properties.Name, attribute.Name) == PropertyGroup(properties.Name, name));

    private string ImportTextSnapshot(XElement source, string part) {
        string family = TextFamily(source);
        CheckDefaults(family);
        XElement? graphic = source.Ancestors().FirstOrDefault(element => element.Name.Namespace == OdfNamespaces.Draw &&
            element.Name.LocalName is not ("text-box" or "a" or "g" or "page"));
        if (graphic == null) throw new NotSupportedException("Drawing import cannot capture text outside an enclosing graphic.");
        string? bound = (string?)source.Attribute(OdfNamespaces.Text + "style-name");
        string? imported = bound == null ? null : ImportStyle(bound, family, part);
        OdfStyle[] chain = OdfTextStyleResolver.Resolve(_source.Styles, source, graphic, part).ToArray();
        if (chain.Any(style => style.Element.Elements(OdfNamespaces.Style + "map").Any()))
            throw new NotSupportedException("Drawing import cannot snapshot conditional text styles.");
        var snapshot = new XElement(OdfNamespaces.Style + "style", new XAttribute(OdfNamespaces.Style + "family", family));
        // Preserve explicit bindings; only source fallbacks absent from the explicit cascade are captured.
        if (bound != null) {
            OdfStyle originalStyle = _source.Styles.FindInPart(family == "paragraph" ? OdfStyleFamily.Paragraph : OdfStyleFamily.Text, bound, part)!;
            if (originalStyle.IsAutomatic) {
                snapshot = new XElement(originalStyle.Element); snapshot.Attribute(OdfNamespaces.Style + "name")?.Remove();
                CaptureLeaderTextOrigins(snapshot, originalStyle.PartPath, commonOnly: false);
                foreach (XAttribute font in snapshot.DescendantsAndSelf().Attributes().Where(attribute => IsFontReference(attribute.Name))) _fontOrigins[font] = originalStyle.PartPath;
            }
            string? parent = originalStyle.IsAutomatic ? originalStyle.ParentStyleName : bound;
            if (!string.IsNullOrEmpty(parent)) snapshot.SetAttributeValue(OdfNamespaces.Style + "parent-style-name", parent);
        }
        foreach (OdfStyle style in chain.Where(candidate => candidate.Element.Name == OdfNamespaces.Style + "default-style")) {
            foreach (XElement properties in style.Element.Elements().Where(element => element.Name == OdfNamespaces.Style + "text-properties" || family == "paragraph" && element.Name == OdfNamespaces.Style + "paragraph-properties")) {
                XElement? target = snapshot.Element(properties.Name);
                if (target == null) { target = new XElement(properties.Name); snapshot.Add(target); }
                CaptureEdgeFallbacks(target, properties, chain.Where(candidate => candidate.Element.Name != OdfNamespaces.Style + "default-style")
                    .Select(candidate => candidate.Element.Element(properties.Name)).OfType<XElement>());
                var existingGroups = new HashSet<string>(target.Attributes().Select(attribute => PropertyGroup(properties.Name, attribute.Name)), StringComparer.Ordinal);
                foreach (XAttribute attribute in properties.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)) {
                    if (IsEdgeProperty(attribute.Name) || existingGroups.Contains(PropertyGroup(properties.Name, attribute.Name)) || chain.Any(candidate => candidate.Element.Name != OdfNamespaces.Style + "default-style" &&
                        candidate.Element.Element(properties.Name) is XElement explicitProperties && HasTextProperty(explicitProperties, attribute.Name))) continue;
                    var copy = new XAttribute(attribute); target.Add(copy);
                    if (IsFontReference(copy.Name)) _fontOrigins.Add(copy, style.PartPath);
                }
                foreach (XElement child in properties.Elements()) if (target.Element(child.Name) == null && !chain.Any(candidate =>
                    candidate.Element.Name != OdfNamespaces.Style + "default-style" && candidate.Element.Element(properties.Name)?.Element(child.Name) != null)) {
                    var copy = new XElement(child); target.Add(copy);
                    CaptureLeaderTextOrigins(copy, style.PartPath, commonOnly: true);
                }
            }
        }
        CaptureRelativeFontSizes(snapshot, chain);
        // Rewrite dependency names before interning snapshots; font declarations may be part-local.
        Rewrite(snapshot, part); _resources.Rewrite(snapshot);
        string key = part + "\0" + imported + "\0" + snapshot.ToString(SaveOptions.DisableFormatting);
        if (_textSnapshots.TryGetValue(key, out string? existing)) return existing;
        string name = ReserveName(); snapshot.SetAttributeValue(OdfNamespaces.Style + "name", name);
        _textSnapshots.Add(key, name); _copies.Add((part, OdfNamespaces.Office + "automatic-styles", snapshot));
        return name;
    }

    private static void CaptureRelativeFontSizes(XElement snapshot, OdfStyle[] chain) {
        foreach (XName attribute in new[] { OdfNamespaces.Fo + "font-size", OdfNamespaces.Style + "font-size-asian", OdfNamespaces.Style + "font-size-complex" }) {
            string? first = chain.Select(style => ((string?)style.TextProperties?.Attribute(attribute))?.Trim()).FirstOrDefault(value => value != null);
            if (first == null || !first.EndsWith("%", StringComparison.Ordinal)) continue;
            string deltaName = attribute == OdfNamespaces.Fo + "font-size" ? "font-size-rel" : "font-size-rel-" + attribute.LocalName.Substring("font-size-".Length);
            if (chain.Select(style => (string?)style.TextProperties?.Attribute(OdfNamespaces.Style + deltaName)).Any(value => value != null &&
                (!OdfLength.Parse(value!).TryToPoints(out double delta) || delta != 0)))
                throw new NotSupportedException("Drawing import cannot snapshot relative font-size changes combined with percentage sizes.");
            OdfStyle? absoluteBase = chain.FirstOrDefault(style => style.TextProperties?.Attribute(attribute) is XAttribute size && !size.Value.Trim().EndsWith("%", StringComparison.Ordinal));
            // Explicit bases travel with the imported style graph. Keep their relative binding and native interpretation.
            if (absoluteBase != null && absoluteBase.Element.Name != OdfNamespaces.Style + "default-style") continue;
            XElement? properties = snapshot.Element(OdfNamespaces.Style + "text-properties");
            if (properties == null) { properties = new XElement(OdfNamespaces.Style + "text-properties"); snapshot.Add(properties); }
            properties.SetAttributeValue(attribute, OdfTextStyleResolver.ResolveFontSize(chain, sizeAttribute: attribute).ToString("R", CultureInfo.InvariantCulture) + "pt");
        }
    }
}
