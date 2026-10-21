using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    // Font references are CSS values even in standalone SVG. Quote embedded
    // names when necessary, and scope them for inline pages sharing a document.
    private static string ScopeEmbeddedFontFamilies(string svg, OfficeFontFaceCollection fonts,
        string prefix, CancellationToken cancellationToken) {
        if (fonts.Faces.Count == 0) return svg;
        var families = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var face in fonts.Faces) {
            families[face.ResourceFamilyName] = prefix + face.ResourceFamilyName;
            families[face.FamilyName] = prefix + face.FamilyName;
        }
        // Parse only markup produced above by our renderer, never external SVG.
        var root = XElement.Parse(svg, LoadOptions.PreserveWhitespace);
        foreach (var element in root.DescendantsAndSelf()) {
            cancellationToken.ThrowIfCancellationRequested();
            var attribute = element.Attribute("font-family");
            if (attribute == null) continue;
            var names = OfficeFontFamilyParser.ParseCssReferences(attribute.Value);
            if (!names.Any(name => families.ContainsKey(name.Name))) continue;
            attribute.Value = string.Join(", ", names.Select(name => {
                if (!name.IsQuoted && IsGenericCssFamily(name.Name)) return name.Name;
                return prefix.Length > 0 && families.TryGetValue(name.Name, out var scoped)
                    ? QuoteCssFamily(scoped)
                    : name.IsQuoted || RequiresQuotedCssFamily(name.Name) ? QuoteCssFamily(name.Name) : name.Name;
            }));
        }
        cancellationToken.ThrowIfCancellationRequested();
        return root.ToString(SaveOptions.DisableFormatting);
    }

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
        value.Equals("emoji", StringComparison.OrdinalIgnoreCase) ||
        value.Equals("fangsong", StringComparison.OrdinalIgnoreCase);

    private static string QuoteCssFamily(string family) => "\"" + EscapeCssString(family) + "\"";
}
