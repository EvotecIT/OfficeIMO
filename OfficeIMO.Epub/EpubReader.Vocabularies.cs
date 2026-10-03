namespace OfficeIMO.Epub;

internal static partial class EpubReader {
    private static void ReadVocabularyPrefixes(XElement? element, EpubPackage package, EpubDiagnosticCollector diagnostics) {
        if (element == null) return;
        var prefixes = EpubVocabulary.ReadPrefixes(GetUnqualifiedAttribute(element, "prefix"), () =>
            diagnostics.Warning("epub.package.prefix-invalid", "Ignored invalid package vocabulary prefix declaration.", package.OpfPath));
        foreach (var pair in prefixes) package.VocabularyPrefixes[pair.Key] = pair.Value;
    }

    private static bool IsRenditionProperty(EpubPackage package, string value, string localName) {
        int colon = value.IndexOf(':');
        return colon > 0 && value.Substring(colon + 1) == localName &&
            package.VocabularyPrefixes.TryGetValue(value.Substring(0, colon), out string? vocabulary) &&
            vocabulary == "http://www.idpf.org/vocab/rendition/#";
    }
}
