using DocumentFormat.OpenXml;

namespace OfficeIMO.Word;

/// <summary>Selects the first markup-compatibility branch whose required namespaces OfficeIMO understands.</summary>
internal static class WordAlternateContentResolver {
    internal const string WordprocessingShapeNamespaceUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

    private static readonly IReadOnlyDictionary<string, string> SupportedPrefixNamespaces =
        new Dictionary<string, string>(System.StringComparer.Ordinal) {
            ["wpc"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
            ["wp14"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
            ["wpg"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
            ["wpi"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
            ["wps"] = WordprocessingShapeNamespaceUri,
            ["w14"] = "http://schemas.microsoft.com/office/word/2010/wordml",
            ["w15"] = "http://schemas.microsoft.com/office/word/2012/wordml",
            ["w16se"] = "http://schemas.microsoft.com/office/word/2015/wordml/symex",
            ["w16cid"] = "http://schemas.microsoft.com/office/word/2016/wordml/cid",
            ["w16"] = "http://schemas.microsoft.com/office/word/2018/wordml",
            ["w16cex"] = "http://schemas.microsoft.com/office/word/2018/wordml/cex",
            ["w16sdtdh"] = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
        };

    private static readonly HashSet<string> SupportedNamespaces = new(System.StringComparer.Ordinal) {
        "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        WordprocessingShapeNamespaceUri,
        "http://schemas.microsoft.com/office/word/2010/wordml",
        "http://schemas.microsoft.com/office/word/2012/wordml",
        "http://schemas.microsoft.com/office/word/2015/wordml/symex",
        "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        "http://schemas.microsoft.com/office/word/2018/wordml",
        "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
    };

    /// <summary>Returns the first supported choice, or the fallback when no choice can be interpreted.</summary>
    internal static OpenXmlCompositeElement? SelectBranch(AlternateContent alternateContent) =>
        (OpenXmlCompositeElement?)SelectBranch(OfficeXmlNode.FromSdk(alternateContent)).SdkElement;

    /// <summary>Returns the first supported <c>mc:Choice</c>, or the <c>mc:Fallback</c> when no choice can be interpreted.</summary>
    internal static OfficeXmlNode SelectBranch(OfficeXmlNode alternateContent) {
        const int Mc = OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.MarkupCompatibility;
        foreach (OfficeXmlNode choice in alternateContent.Children(Mc, "Choice")) {
            if (AreRequiredNamespacesSupported(choice)) return choice;
        }
        return alternateContent.Child(Mc, "Fallback");
    }

    /// <summary>Enumerates immediate content, flattening only selected compatibility wrappers.</summary>
    internal static IEnumerable<OpenXmlElement> EnumerateEffectiveChildren(OpenXmlElement parent) {
        foreach (OpenXmlElement child in parent.ChildElements) {
            if (child is AlternateContent alternateContent) {
                OpenXmlCompositeElement? branch = SelectBranch(alternateContent);
                if (branch != null) {
                    foreach (OpenXmlElement selectedChild in EnumerateEffectiveChildren(branch)) yield return selectedChild;
                }
            } else yield return child;
        }
    }

    private static bool AreRequiredNamespacesSupported(OfficeXmlNode choice) {
        string? requires = choice.Attribute("Requires");
        if (string.IsNullOrWhiteSpace(requires)) return false;
        string[] prefixes = requires!.Split(new[] { ' ', '\t', '\r', '\n' }, System.StringSplitOptions.RemoveEmptyEntries);
        return prefixes.Length > 0 && prefixes.All(prefix => {
            string? namespaceUri = OfficeIMO.Word.OpenXml.WordXmlNodes.LookupNamespace(choice, prefix);
            // Older OfficeIMO documents could omit the canonical declaration from header/footer parts.
            if (namespaceUri == null) SupportedPrefixNamespaces.TryGetValue(prefix, out namespaceUri);
            return namespaceUri != null && SupportedNamespaces.Contains(namespaceUri);
        });
    }
}
