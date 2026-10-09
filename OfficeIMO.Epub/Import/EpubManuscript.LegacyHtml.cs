using System.Globalization;

namespace OfficeIMO.Epub;

public static partial class EpubManuscript {
    private static void NormalizeLegacyHtml(XElement element, List<OfficeConversionFidelityDiagnostic> diagnostics) {
        if (element.Name.Namespace != Xhtml) return;
        string name = element.Name.LocalName;
        if (name is "table" or "tr" or "td" or "th" or "div" or "p" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "hr") {
            Translate("align", value => Alignment(name, value));
        }
        if (name is "table" or "tr" or "td" or "th")
            Translate("valign", value => value is "top" or "middle" or "bottom" or "baseline" ? "vertical-align:" + value : null);
        if (name is "table" or "td" or "th" or "hr") {
            Translate("width", value => Length(value) is string length ? "width:" + length : null);
            Translate("height", value => Length(value) is string length ? "height:" + length : null);
        }
        if (name == "ul") Translate("type", value => value is "disc" or "circle" or "square" ? "list-style-type:" + value : null);
        if (name == "table") {
            Translate("cellspacing", value => Length(value, false) is string length ? "border-spacing:" + length + ";border-collapse:separate" : null);
            Translate("border", value => Length(value, false) is string length ? "border:" + length + " solid" : null);
            if (element.Attribute("cellpadding") is XAttribute padding) {
                string? length = Length(padding.Value, false);
                if (length != null)
                    foreach (XElement cell in element.Descendants().Where(child => child.Name == Xhtml + "td" || child.Name == Xhtml + "th")
                        .Where(child => child.Ancestors(Xhtml + "table").First() == element)) PrependStyle(cell, "padding:" + length);
                Record(padding, length != null);
            }
            if (element.Attribute("summary") is XAttribute summary) {
                if (element.Attribute("aria-description") == null) element.SetAttributeValue("aria-description", summary.Value);
                Record(summary, true);
            }
        }
        // Older documentation generators use definition lists for navigation with
        // terms but no descriptions. Preserve their links and nested lists as a list.
        if (name == "dl" && element.Elements().Any() && element.Elements().All(child => child.Name == Xhtml + "dt")) {
            element.Name = Xhtml + "ul";
            foreach (XElement term in element.Elements()) term.Name = Xhtml + "li";
            AddDiagnostic(diagnostics, "EPUB_IMPORT_LEGACY_LIST_NORMALIZED", "A term-only HTML definition list was represented as an unordered list.", "dl", OfficeConversionLossKind.Approximation);
        } else if (name == "dl" && element.Elements().LastOrDefault()?.Name == Xhtml + "dt") {
            element.Add(new XElement(Xhtml + "dd"));
            AddDiagnostic(diagnostics, "EPUB_IMPORT_LEGACY_LIST_NORMALIZED", "A trailing term without a definition received an empty description.", "dl", OfficeConversionLossKind.Approximation);
        }

        void Translate(string attributeName, Func<string, string?> translate) {
            if (element.Attribute(attributeName) is not XAttribute attribute) return;
            string? css = translate(attribute.Value.Trim().ToLowerInvariant());
            if (css != null) PrependStyle(element, css);
            Record(attribute, css != null);
        }
        void Record(XAttribute attribute, bool represented) {
            AddDiagnostic(diagnostics, represented ? "EPUB_IMPORT_LEGACY_ATTRIBUTE_NORMALIZED" : "EPUB_IMPORT_LEGACY_ATTRIBUTE_OMITTED",
                represented ? "An obsolete HTML attribute was represented through CSS or accessibility metadata." : "An unsupported obsolete HTML attribute value was omitted.",
                name + "@" + attribute.Name.LocalName, represented ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.Omission);
            attribute.Remove();
        }
    }

    private static string? Alignment(string name, string value) {
        if (name == "table") return value switch {
            "center" => "margin-left:auto;margin-right:auto", "left" or "right" => "float:" + value, _ => null
        };
        if (name == "hr") return value switch {
            "center" => "margin-left:auto;margin-right:auto", "left" => "margin-left:0;margin-right:auto",
            "right" => "margin-left:auto;margin-right:0", _ => null
        };
        return value is "left" or "right" or "center" or "justify" ? "text-align:" + value : null;
    }

    private static string? Length(string value, bool allowPercent = true) {
        value = value.Trim();
        bool percent = allowPercent && value.EndsWith("%", StringComparison.Ordinal);
        string numeric = percent ? value.Substring(0, value.Length - 1) : value;
        return uint.TryParse(numeric, NumberStyles.None, CultureInfo.InvariantCulture, out uint number)
            ? number.ToString(CultureInfo.InvariantCulture) + (percent ? "%" : "px") : null;
    }

    private static void PrependStyle(XElement element, string declaration) =>
        element.SetAttributeValue("style", declaration + ";" + ((string?)element.Attribute("style") ?? string.Empty));
}
