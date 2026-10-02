namespace OfficeIMO.Epub;

internal static partial class EpubReader {
    private static void ReadVocabularyPrefixes(XElement? element, EpubPackage package, EpubDiagnosticCollector diagnostics) {
        if (element == null) return;
        string[] parts = GetUnqualifiedAttribute(element, "prefix")
            .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        for (int index = 0; index < parts.Length; index += 2) {
            string prefix = parts[index];
            if (index + 1 >= parts.Length || !prefix.EndsWith(":", StringComparison.Ordinal) || prefix.Length < 2 ||
                !Uri.TryCreate(parts[index + 1], UriKind.Absolute, out _)) {
                diagnostics.Warning("epub.package.prefix-invalid", "Ignored invalid package vocabulary prefix declaration.", package.OpfPath);
                continue;
            }
            package.VocabularyPrefixes[prefix.Substring(0, prefix.Length - 1)] = parts[index + 1];
        }
    }

    private static bool IsRenditionProperty(EpubPackage package, string value, string localName) {
        int colon = value.IndexOf(':');
        return colon > 0 && value.Substring(colon + 1) == localName &&
            package.VocabularyPrefixes.TryGetValue(value.Substring(0, colon), out string? vocabulary) &&
            vocabulary == "http://www.idpf.org/vocab/rendition/#";
    }
}
