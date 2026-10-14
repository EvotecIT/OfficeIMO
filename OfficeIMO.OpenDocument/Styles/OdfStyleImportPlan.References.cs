namespace OfficeIMO.OpenDocument;

internal sealed partial class OdfStyleImportPlan {
    private readonly Dictionary<XAttribute, (string Part, bool CommonOnly)> _leaderTextOrigins = new();

    /// <summary>Retains a copied leader binding's source scope before its definition is detached or synthesized.</summary>
    private void CaptureLeaderTextOrigins(XElement copy, string part, bool commonOnly) {
        foreach (XAttribute attribute in copy.DescendantsAndSelf().Attributes(OdfNamespaces.Style + "leader-text-style"))
            _leaderTextOrigins[attribute] = (part, commonOnly);
    }

    internal void Rewrite(XElement root, string part) {
        foreach (XElement element in root.DescendantsAndSelf().ToArray()) {
            XElement? definition = element.AncestorsAndSelf().FirstOrDefault(candidate => candidate.Name == OdfNamespaces.Style + "style" || candidate.Name.Namespace == OdfNamespaces.Number);
            string? family = (string?)definition?.Attribute(OdfNamespaces.Style + "family");
            foreach (XAttribute attribute in element.Attributes().ToArray()) {
                string value = attribute.Value; XName name = attribute.Name;
                if (value.Length == 0 && name.Namespace == OdfNamespaces.Draw && name.LocalName is "marker-start" or "marker-end" or "fill-gradient-name" or "opacity-name" or "stroke-dash" or "fill-image-name") continue;
                if (name == OdfNamespaces.Draw + "style-name") attribute.Value = ImportStyle(value, DrawingFamily(element), part);
                else if (name == OdfNamespaces.Draw + "text-style-name") attribute.Value = ImportStyle(value, "paragraph", part);
                else if (name == OdfNamespaces.Draw + "class-names") attribute.Value = ImportStyleList(value, "graphic", part);
                else if (name == OdfNamespaces.Presentation + "style-name") attribute.Value = ImportStyle(value, "presentation", part);
                else if (name == OdfNamespaces.Presentation + "class-names") attribute.Value = ImportStyleList(value, "presentation", part);
                else if (name == OdfNamespaces.Text + "style-name") {
                    if (element.Name == OdfNamespaces.Text + "list" || element.Name == OdfNamespaces.Text + "numbered-paragraph") attribute.Value = ImportNamedDefinition(value, "list", part);
                    else if (!_textSources.ContainsKey(element)) attribute.Value = ImportStyle(value, TextFamily(element), part);
                } else if (name == OdfNamespaces.Text + "class-names") attribute.Value = ImportStyleList(value, TextFamily(element), part);
                else if (name == OdfNamespaces.Text + "visited-style-name") attribute.Value = ImportStyle(value, "text", part);
                else if (name == OdfNamespaces.Style + "parent-style-name" || name == OdfNamespaces.Style + "next-style-name" || name == OdfNamespaces.Style + "apply-style-name") {
                    if (element.Name == OdfNamespaces.Style + "master-page") attribute.Value = RemapMaster(value);
                    else if (definition?.Name.Namespace == OdfNamespaces.Number) attribute.Value = ImportNamedDefinition(value, "data", part, commonOnly: true);
                    else attribute.Value = ImportStyle(value, family ?? throw new NotSupportedException("Imported style reference has no supported family."), part, true);
                } else if (name == OdfNamespaces.Style + "list-style-name") {
                    if (value.Length > 0) attribute.Value = ImportNamedDefinition(value, "list", part);
                } else if (name == OdfNamespaces.Style + "leader-text-style") attribute.Value = RewriteLeaderTextStyle(attribute);
                else if (name == OdfNamespaces.Style + "data-style-name" || name == OdfNamespaces.Style + "percentage-data-style-name") attribute.Value = ImportNamedDefinition(value, "data", part);
                else if (IsFontReference(name)) attribute.Value = ImportFont(value, _fontOrigins.TryGetValue(attribute, out string? origin) ? origin : part);
                else if (name == OdfNamespaces.Style + "master-page-name") attribute.Value = RemapMaster(value);
                else if (name == OdfNamespaces.Draw + "fill-gradient-name") attribute.Value = ImportNamedDefinition(value, "gradient", part);
                else if (name == OdfNamespaces.Draw + "opacity-name") attribute.Value = ImportNamedDefinition(value, "opacity", part);
                else if (name == OdfNamespaces.Draw + "stroke-dash") attribute.Value = ImportNamedDefinition(value, "stroke-dash", part);
                else if (name == OdfNamespaces.Draw + "stroke-dash-names") attribute.Value = string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(item => ImportNamedDefinition(item, "stroke-dash", part)));
                else if (name == OdfNamespaces.Draw + "fill-hatch-name") attribute.Value = ImportNamedDefinition(value, "hatch", part);
                else if (name == OdfNamespaces.Draw + "marker-start" || name == OdfNamespaces.Draw + "marker-end") attribute.Value = ImportNamedDefinition(value, "marker", part);
                else if (name == OdfNamespaces.Draw + "fill-image-name") attribute.Value = ImportNamedDefinition(value, "fill-image", part);
                else if (name == OdfNamespaces.XLink + "href" && value.StartsWith("#", StringComparison.Ordinal) && MatchesKind(root, "gradient"))
                    attribute.Value = "#" + ImportNamedDefinition(Uri.UnescapeDataString(value.Substring(1)), "gradient", part);
                else if (name.Namespace == OdfNamespaces.Table || name.Namespace == OdfNamespaces.Presentation || name.Namespace == OdfNamespaces.Text || name.Namespace == OdfNamespaces.Style) {
                    if (name.LocalName.EndsWith("style-name", StringComparison.Ordinal) && name != OdfNamespaces.Draw + "master-page-name" ||
                        name.Namespace == OdfNamespaces.Presentation && name.LocalName.StartsWith("use-", StringComparison.Ordinal))
                        throw new NotSupportedException("Drawing import does not remap this dependency: " + name);
                }
            }
            if (_textSources.TryGetValue(element, out XElement? original))
                element.SetAttributeValue(OdfNamespaces.Text + "style-name", ImportTextSnapshot(original, part));
            else if (element.Name == OdfNamespaces.Draw + "page" || element.Name == OdfNamespaces.Style + "master-page" ||
                element.Name.Namespace == OdfNamespaces.Draw && element.Name.LocalName is "rect" or "ellipse" or "circle" or "line" or "path" or "polygon" or "polyline" or "connector" or "frame" or "custom-shape" or "caption" or "measure" or "regular-polygon") {
                if (element.Attribute(OdfNamespaces.Draw + "style-name") == null)
                    element.SetAttributeValue(OdfNamespaces.Draw + "style-name", ImportDefault(DrawingFamily(element), part));
            } else if ((element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h") && element.Attribute(OdfNamespaces.Text + "style-name") == null)
                element.SetAttributeValue(OdfNamespaces.Text + "style-name", ImportDefault("paragraph", part));
        }
    }
    // Keep origin tuples and validation out of the recursive dependency-rewrite frame.
    private string RewriteLeaderTextStyle(XAttribute attribute) {
        XElement? stop = attribute.Parent, stops = stop?.Parent, properties = stops?.Parent, definition = properties?.Parent;
        if (stop?.Name != OdfNamespaces.Style + "tab-stop" || stops?.Name != OdfNamespaces.Style + "tab-stops" ||
            properties?.Name != OdfNamespaces.Style + "paragraph-properties" ||
            definition?.Name != OdfNamespaces.Style + "style" && definition?.Name != OdfNamespaces.Style + "default-style" ||
            !_leaderTextOrigins.TryGetValue(attribute, out var origin))
            throw new NotSupportedException("Drawing import cannot resolve this leader text style in its original paragraph-property scope.");
        return ImportStyle(attribute.Value, "text", origin.Part, origin.CommonOnly);
    }
    private string RemapMaster(string name) => name == _sourceMaster ? _destinationMaster! : throw new NotSupportedException("Drawing import does not copy other master dependencies: " + name);
    private string ImportStyleList(string value, string family, string part) => string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(name => ImportStyle(name, family, part)));
    private static string DrawingFamily(XElement element) => element.Name == OdfNamespaces.Draw + "page" || element.Name == OdfNamespaces.Style + "master-page" ? "drawing-page" : "graphic";
    private static string TextFamily(XElement element) => element.Name.LocalName switch {
        "p" or "h" => "paragraph", "span" or "a" or "list-level-style-number" or "list-level-style-bullet" or "list-level-style-image" => "text",
        _ => throw new NotSupportedException("Drawing import does not remap this text style family: " + element.Name.LocalName)
    };
}
