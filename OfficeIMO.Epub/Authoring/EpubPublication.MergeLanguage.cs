namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static bool IsMergeLanguageAttribute(XAttribute attribute) =>
        attribute.Name == "lang" || attribute.Name == XNamespace.Xml + "lang" || attribute.Name == "dir";

    private static bool SameMergeScaffoldAttributes(XElement first, XElement second, bool preserveLanguage) {
        if (!preserveLanguage) return SameMergeAttributes(first, second);
        var left = new XElement(first.Name, first.Attributes().Where(attribute => !IsMergeLanguageAttribute(attribute)));
        var right = new XElement(second.Name, second.Attributes().Where(attribute => !IsMergeLanguageAttribute(attribute)));
        return SameMergeAttributes(left, right);
    }

    private static XElement CreateMergeLanguageWrapper(XElement firstRoot, XElement secondRoot) {
        foreach (XElement root in new[] { firstRoot, secondRoot }) {
            foreach (XElement element in new[] { root, root.Element(Html + "body")! }) {
                string? direction = (string?)element.Attribute("dir");
                if (direction != null && !string.Equals(direction, "ltr", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(direction, "rtl", StringComparison.OrdinalIgnoreCase))
                    throw new NotSupportedException("Language-preserving chapter merges require explicit ltr/rtl or absent root/body direction.");
                string? htmlLanguage = (string?)element.Attribute("lang"), xmlLanguage = (string?)element.Attribute(XNamespace.Xml + "lang");
                if (htmlLanguage != null && xmlLanguage != null && !string.Equals(htmlLanguage, xmlLanguage, StringComparison.OrdinalIgnoreCase))
                    throw new NotSupportedException("Resolve conflicting lang and xml:lang before merging chapter language contexts.");
            }
        }
        XElement body = secondRoot.Element(Html + "body")!;
        string language = (string?)body.Attribute(XNamespace.Xml + "lang") ?? (string?)body.Attribute("lang") ??
            (string?)secondRoot.Attribute(XNamespace.Xml + "lang") ?? (string?)secondRoot.Attribute("lang") ?? string.Empty;
        string directionValue = (string?)body.Attribute("dir") ?? (string?)secondRoot.Attribute("dir") ?? "ltr";
        return new XElement(Html + "div", new XAttribute("lang", language), new XAttribute(XNamespace.Xml + "lang", language),
            new XAttribute("dir", directionValue.ToLowerInvariant()));
    }
}
