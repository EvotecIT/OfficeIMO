using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

/// <summary>Rejects invalid rendering vocabulary before it can reach HTML generation.</summary>
internal static class CslStyleValidation {
    private static readonly HashSet<string> Elements = new HashSet<string>(new[] {
        "style", "info", "locale", "macro", "citation", "bibliography", "sort", "key", "layout",
        "text", "number", "label", "names", "name", "name-part", "et-al", "substitute", "date", "date-part",
        "group", "choose", "if", "else-if", "else", "terms", "term", "single", "multiple", "style-options"
    }, StringComparer.Ordinal);
    private static readonly IReadOnlyDictionary<string, string[]> Values = new Dictionary<string, string[]>(StringComparer.Ordinal) {
        ["font-style"] = new[] { "normal", "italic", "oblique" },
        ["font-variant"] = new[] { "normal", "small-caps" },
        ["font-weight"] = new[] { "normal", "bold", "light" },
        ["text-decoration"] = new[] { "none", "underline" },
        ["vertical-align"] = new[] { "baseline", "sup", "sub" },
        ["text-case"] = new[] { "lowercase", "uppercase", "capitalize-first", "capitalize-all", "sentence", "title" },
        ["display"] = new[] { "block", "left-margin", "right-inline", "indent" },
        ["second-field-align"] = new[] { "flush", "margin" },
        ["hanging-indent"] = new[] { "true", "false" },
        ["match"] = new[] { "all", "any", "none", "last-digit", "last-two-digits", "whole-number" },
        ["page-range-format"] = new[] { "expanded", "minimal", "minimal-two", "chicago", "chicago-15", "chicago-16" },
        ["demote-non-dropping-particle"] = new[] { "never", "sort-only", "display-and-sort" },
        ["givenname-disambiguation-rule"] = new[] { "all-names", "all-names-with-initials", "primary-name", "primary-name-with-initials", "by-cite" },
        ["collapse"] = new[] { "citation-number", "year", "year-suffix", "year-suffix-ranged" },
        ["subsequent-author-substitute-rule"] = new[] { "complete-all", "complete-each", "partial-each", "partial-first" }
    };

    internal static void Validate(XElement root, CancellationToken token = default) {
        if ((string?)root.Attribute("class") != "note" && (string?)root.Attribute("class") != "in-text")
            throw new InvalidDataException("CSL style class must be note or in-text.");
        ValidateElements(root, token);
        foreach (XElement section in root.Elements().Where(element => element.Name.LocalName == "citation" || element.Name.LocalName == "bibliography"))
            if (section.Elements(CslStyle.Namespace + "layout").Count() != 1) throw new InvalidDataException("A CSL rendering section requires exactly one layout.");
    }

    internal static void ValidateLocale(XElement root, CancellationToken token) => ValidateElements(root, token);

    private static void ValidateElements(XElement root, CancellationToken token) {
        foreach (XElement element in root.DescendantsAndSelf()) {
            token.ThrowIfCancellationRequested();
            // CSL metadata admits Dublin Core and other foreign elements.
            if (element.AncestorsAndSelf().Any(ancestor => ancestor.Name == CslStyle.Namespace + "info")) continue;
            if (element.Name.Namespace != CslStyle.Namespace || !Elements.Contains(element.Name.LocalName))
                throw new InvalidDataException("Unknown CSL element '" + element.Name + "'.");
            foreach (XAttribute attribute in element.Attributes()) {
                if (attribute.IsNamespaceDeclaration || attribute.Name.Namespace == XNamespace.Xml) continue;
                if (attribute.Name.Namespace != XNamespace.None) throw new InvalidDataException("Foreign attributes are not allowed on CSL rendering elements.");
                if (Values.TryGetValue(attribute.Name.LocalName, out string[]? allowed) && !allowed.Contains(attribute.Value, StringComparer.Ordinal))
                    throw new InvalidDataException("Invalid CSL " + attribute.Name.LocalName + " value '" + attribute.Value + "'.");
                if (attribute.Name.LocalName == "line-spacing" || attribute.Name.LocalName == "entry-spacing") {
                    int minimum = attribute.Name.LocalName == "line-spacing" ? 1 : 0;
                    if (!int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int spacing) || spacing < minimum)
                        throw new InvalidDataException("Invalid CSL " + attribute.Name.LocalName + ": an integer of at least " + minimum + " is required.");
                }
            }
        }
    }
}
