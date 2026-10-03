using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal readonly partial struct CslText {
    /// <summary>Resolves input emphasis against style formatting and removes input-only case protection.</summary>
    internal string FinalizeHtml(int maximumCharacters, CancellationToken token, out int maximumLeftMarginCharacters) {
        XElement root = CslStyle.ReadXml("<root>" + Html + "</root>", (int)Math.Min(int.MaxValue, (long)maximumCharacters + 13), 1024, token);
        maximumLeftMarginCharacters = MeasureLeftMarginCharacters(root, token);
        var output = new StringBuilder(Html.Length);
        AppendMarkup(root.Nodes(), new Dictionary<string, string>(StringComparer.Ordinal), output, token);
        return output.ToString();
    }

    internal int MeasureLeftMarginCharacters(int maximumCharacters, CancellationToken token) {
        if (!Html.Contains("csl-left-margin")) return 0;
        XElement root = CslStyle.ReadXml("<root>" + Html + "</root>", (int)Math.Min(int.MaxValue, (long)maximumCharacters + 13), 1024, token);
        return MeasureLeftMarginCharacters(root, token);
    }

    private static int MeasureLeftMarginCharacters(XElement root, CancellationToken token) {
        int maximum = 0;
        foreach (XElement element in root.Descendants()) {
            token.ThrowIfCancellationRequested();
            if (element.Name.LocalName == "div" && (string?)element.Attribute("class") == "csl-left-margin")
                maximum = Math.Max(maximum, element.Value.Length);
        }
        return maximum;
    }

    private static void AppendMarkup(IEnumerable<XNode> nodes, IReadOnlyDictionary<string, string> inherited, StringBuilder output, CancellationToken token) {
        foreach (XNode node in nodes) {
            token.ThrowIfCancellationRequested();
            if (node is XText text) { output.Append(Escape(text.Value)); continue; }
            if (node is not XElement element) continue;
            if (element.Name.LocalName != "div" && !element.DescendantNodes().OfType<XText>().Any(text => text.Value.Length > 0)) continue;
            var current = inherited.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
            string name = element.Name.LocalName;
            string? inputClass = (string?)element.Attribute("class");
            bool nodecor = inputClass == "nodecor" || inputClass == "nocase nodecor";
            // Input protection establishes the parent context for this span's
            // own formatting, just as an enclosing nodecor span would do.
            if (nodecor)
                foreach (var property in inherited) current[property.Key] = DefaultProperty(property.Key);
            var properties = new Dictionary<string, string>(StringComparer.Ordinal);
            if (name == "i") properties["font-style"] = "italic";
            else if (name == "b") properties["font-weight"] = "bold";
            else if (name == "sup" || name == "sub") properties["vertical-align"] = name;
            else if ((string?)element.Attribute("style") is string css)
                foreach (string declaration in css.Split(';')) {
                    int colon = declaration.IndexOf(':');
                    if (colon > 0) properties[declaration.Substring(0, colon)] = declaration.Substring(colon + 1);
                }
            if ((string?)element.Attribute("data-csl-flip") == "true") {
                foreach (string property in properties.Keys.ToArray())
                    if (current.TryGetValue(property, out string? parent) && parent == properties[property]) properties[property] = DefaultProperty(property);
            }
            if (nodecor)
                foreach (var property in inherited)
                    if (!properties.ContainsKey(property.Key) && property.Value != DefaultProperty(property.Key)) properties[property.Key] = DefaultProperty(property.Key);
            foreach (var property in properties) current[property.Key] = property.Value;

            string? opening = null, closing = null;
            if (properties.Count == 1 && properties.TryGetValue("font-style", out string? italic) && italic == "italic") { opening = "<i>"; closing = "</i>"; }
            else if (properties.Count == 1 && properties.TryGetValue("font-weight", out string? bold) && bold == "bold") { opening = "<b>"; closing = "</b>"; }
            else if (properties.Count == 1 && properties.TryGetValue("vertical-align", out string? vertical) && (vertical == "sup" || vertical == "sub")) { opening = "<" + vertical + ">"; closing = "</" + vertical + ">"; }
            else if (properties.Count > 0) {
                opening = "<span style=\"" + string.Concat(properties.Select(pair => Escape(pair.Key) + ":" + Escape(pair.Value) + ";")) + "\">";
                closing = "</span>";
            } else if (name == "div") { opening = "<div class=\"" + Escape((string?)element.Attribute("class") ?? string.Empty) + "\">"; closing = "</div>"; }
            else if (name == "a" && (string?)element.Attribute("href") is string href) { opening = "<a href=\"" + Escape(href) + "\">"; closing = "</a>"; }
            if (opening != null) output.Append(opening);
            AppendMarkup(element.Nodes(), current, output, token);
            if (closing != null) output.Append(closing);
        }
    }

    private static string DefaultProperty(string name) => name == "vertical-align" ? "baseline" : name == "text-decoration" ? "none" : "normal";
}
