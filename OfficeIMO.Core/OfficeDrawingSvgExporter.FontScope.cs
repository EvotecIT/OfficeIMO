using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    // Font references are CSS values even in standalone SVG. Quote embedded
    // names when necessary, and scope them for inline pages sharing a document.
    private static string ScopeEmbeddedFontFamilies(string svg, OfficeFontFaceCollection fonts,
        string prefix, CancellationToken cancellationToken) {
        if (fonts.Faces.Count == 0 || !ContainsFontFamilyAttribute(svg)) return svg;
        var families = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var face in fonts.Faces) {
            if (!face.AllowsOutlineEmbedding) continue;
            families[face.ResourceFamilyName] = prefix + face.ResourceFamilyName;
            families[face.FamilyName] = prefix + face.FamilyName;
        }
        // Parse only markup produced above by our renderer, never external SVG.
        var root = XElement.Parse(svg, LoadOptions.PreserveWhitespace);
        bool changed = false;
        foreach (var element in root.DescendantsAndSelf()) {
            cancellationToken.ThrowIfCancellationRequested();
            var attribute = element.Attribute("font-family");
            if (attribute == null) continue;
            var names = OfficeFontFamilyParser.ParseCssReferences(attribute.Value);
            if (!names.Any(name => families.ContainsKey(name.Name))) continue;
            string reference = string.Join(", ", names.Select(name => {
                if (!name.IsQuoted && IsGenericCssFamily(name.Name)) return name.Name;
                return prefix.Length > 0 && families.TryGetValue(name.Name, out var scoped)
                    ? QuoteCssFamily(scoped)
                    : name.IsQuoted || RequiresQuotedCssFamily(name.Name) ? QuoteCssFamily(name.Name) : name.Name;
            }));
            if (string.Equals(attribute.Value, reference, StringComparison.Ordinal)) continue;
            attribute.Value = reference;
            changed = true;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return changed ? root.ToString(SaveOptions.DisableFormatting) : svg;
    }

    // An XML attribute is its name, optional white space and '='. Style sheet declarations
    // use ':' instead, so markup without such an attribute needs no parse.
    private static bool ContainsFontFamilyAttribute(string markup) {
        const string name = "font-family";
        for (int index = markup.IndexOf(name, StringComparison.Ordinal); index >= 0;
            index = markup.IndexOf(name, index + name.Length, StringComparison.Ordinal)) {
            int next = index + name.Length;
            while (next < markup.Length && IsXmlWhiteSpace(markup[next])) next++;
            if (next < markup.Length && markup[next] == '=') return true;
        }
        return false;
    }

    // Only start tags can introduce a text or tspan element; '<' is escaped everywhere else
    // in well-formed markup, so a fragment without such a tag has no text families to collect.
    private static bool ContainsSvgTextElement(string markup) {
        for (int index = markup.IndexOf('<'); index >= 0; index = markup.IndexOf('<', index + 1)) {
            int start = index + 1, end = start;
            while (end < markup.Length && !IsXmlWhiteSpace(markup[end]) && markup[end] != '>' && markup[end] != '/') end++;
            int colon = markup.LastIndexOf(':', end - 1, end - start);
            if (colon >= 0) start = colon + 1;
            int length = end - start;
            if (length == 4 && string.CompareOrdinal(markup, start, "text", 0, 4) == 0 ||
                length == 5 && string.CompareOrdinal(markup, start, "tspan", 0, 5) == 0) return true;
        }
        return false;
    }

    private static bool IsXmlWhiteSpace(char character) =>
        character == ' ' || character == '\t' || character == '\r' || character == '\n';

    private static bool RequiresQuotedCssFamily(string family) {
        if (string.IsNullOrWhiteSpace(family)) return true;
        if (IsGenericCssFamily(family)) return false;

        var identifiers = family.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (identifiers.Length == 0) return true;
        if (!string.Equals(family, string.Join(" ", identifiers), StringComparison.Ordinal)) return true;
        for (int index = 0; index < identifiers.Length; index++) {
            if (!IsUnquotedCssIdentifier(identifiers[index]) || IsCssWideKeyword(identifiers[index])) return true;
        }
        return false;
    }

    private static bool IsUnquotedCssIdentifier(string value) {
        if (value.Length == 0) return false;
        int index = 0;
        if (value[0] == '-') {
            index = 1;
            if (index == value.Length) return false;
            if (value[index] == '-') index++;
        }
        if (index < value.Length && !IsCssIdentifierStart(value[index])) return false;
        for (index++; index < value.Length; index++) {
            char character = value[index];
            if (!IsCssIdentifierStart(character) && !char.IsDigit(character) && character != '-') return false;
        }
        return true;
    }

    private static bool IsCssIdentifierStart(char character) =>
        character == '_' || char.IsLetter(character) || character >= 0x80;

    private static bool IsCssWideKeyword(string value) =>
        value.Equals("initial", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("inherit", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("unset", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("revert", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("revert-layer", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("default", StringComparison.OrdinalIgnoreCase);

    private static bool IsGenericCssFamily(string value) =>
        value.Equals("serif", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("sans-serif", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("monospace", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("cursive", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("fantasy", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("system-ui", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("ui-serif", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("ui-sans-serif", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("ui-monospace", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("ui-rounded", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("math", StringComparison.OrdinalIgnoreCase) ||
        IsFunctionalGenericCssFamily(value);

    // CSS Fonts 4 script-specific generics are functions, not literal family names.
    private static bool IsFunctionalGenericCssFamily(string value) {
        if (!value.StartsWith("generic(", StringComparison.OrdinalIgnoreCase) ||
            value[value.Length - 1] != ')') return false;
        string argument = value.Substring(8, value.Length - 9).Trim(' ', '\t', '\r', '\n', '\f');
        return argument.Equals("fangsong", StringComparison.OrdinalIgnoreCase) ||
            argument.Equals("kai", StringComparison.OrdinalIgnoreCase) ||
            argument.Equals("khmer-mul", StringComparison.OrdinalIgnoreCase) ||
            argument.Equals("nastaliq", StringComparison.OrdinalIgnoreCase);
    }

    private static string QuoteCssFamily(string family) => "\"" + EscapeCssString(family) + "\"";
}
