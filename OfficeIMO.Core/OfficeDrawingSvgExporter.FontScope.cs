using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    // Inline SVG styles share the host document's font namespace. Scope both
    // registrations and references, including rich text and nested drawing groups.
    private static string ScopeEmbeddedFontFamilies(string svg, OfficeFontFaceCollection fonts,
        string prefix, CancellationToken cancellationToken) {
        if (prefix.Length == 0 || fonts.Faces.Count == 0) return svg;
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
            var names = OfficeFontFamilyParser.Parse(attribute.Value).ToArray();
            if (!names.Any(families.ContainsKey)) continue;
            attribute.Value = string.Join(", ", names.Select(name => families.TryGetValue(name, out var scoped)
                ? "\"" + EscapeCssString(scoped) + "\"" : name));
        }
        cancellationToken.ThrowIfCancellationRequested();
        return root.ToString(SaveOptions.DisableFormatting);
    }
}
