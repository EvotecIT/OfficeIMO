namespace OfficeIMO.Epub;

/// <summary>Shared package vocabulary expansion for extraction and authoring.</summary>
internal static class EpubVocabulary {
    private const string PackageVocabulary = "http://idpf.org/epub/vocab/package/#";

    // Prefix declarations cannot alias vocabularies that EPUB assigns without prefixes.
    internal static void ValidateDeclaration(string prefix, string vocabularyUri) {
        XmlConvert.VerifyNCName(prefix);
        if (prefix == "_") throw new ArgumentException("EPUB reserves the underscore prefix.", nameof(prefix));
        if (!Uri.TryCreate(vocabularyUri, UriKind.Absolute, out Uri? uri) || !uri.IsWellFormedOriginalString() || vocabularyUri.Any(char.IsWhiteSpace))
            throw new ArgumentException("Vocabulary must have a well-formed absolute URI without whitespace.", nameof(vocabularyUri));
        if (new[] { PackageVocabulary, "http://idpf.org/epub/vocab/package/item/#", "http://idpf.org/epub/vocab/package/itemref/#",
            "http://idpf.org/epub/vocab/package/link/#", "http://idpf.org/epub/vocab/structure/#", "http://purl.org/dc/elements/1.1/" }
            .Contains(uri.AbsoluteUri, StringComparer.Ordinal))
            throw new ArgumentException("Default EPUB vocabularies and Dublin Core elements cannot be assigned a prefix.", nameof(vocabularyUri));
    }

    internal static void ValidatePropertyName(XElement package, string property) {
        int colon = property.IndexOf(':');
        if (property.Any(char.IsWhiteSpace) || colon == 0 || colon == property.Length - 1 ||
            (colon > 0 && !ReadPrefixes((string?)package.Attribute("prefix") ?? string.Empty).ContainsKey(property.Substring(0, colon))))
            throw new ArgumentException("Metadata properties require a nonempty reference and a declared or reserved prefix.", nameof(property));
    }

    internal static Dictionary<string, string> ReadPrefixes(string declaration, Action? invalid = null) {
        var prefixes = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["rendition"] = "http://www.idpf.org/vocab/rendition/#", ["dcterms"] = "http://purl.org/dc/terms/",
            ["schema"] = "http://schema.org/", ["media"] = "http://www.idpf.org/epub/vocab/overlays/#",
            ["a11y"] = "http://www.idpf.org/epub/vocab/package/a11y/#", ["marc"] = "http://id.loc.gov/vocabulary/",
            ["onix"] = "http://www.editeur.org/ONIX/book/codelists/current.html#", ["xsd"] = "http://www.w3.org/2001/XMLSchema#"
        };
        string[] parts = declaration.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        for (int index = 0; index < parts.Length; index += 2) {
            if (index + 1 >= parts.Length || !parts[index].EndsWith(":", StringComparison.Ordinal) || parts[index].Length < 2 ||
                !Uri.TryCreate(parts[index + 1], UriKind.Absolute, out _)) { invalid?.Invoke(); continue; }
            prefixes[parts[index].Substring(0, parts[index].Length - 1)] = parts[index + 1];
        }
        return prefixes;
    }
    internal static string Expand(XElement package, string property) {
        int colon = property.IndexOf(':');
        if (colon < 0) return PackageVocabulary + property;
        var prefixes = ReadPrefixes((string?)package.Attribute("prefix") ?? string.Empty);
        return prefixes.TryGetValue(property.Substring(0, colon), out string? vocabulary) ? vocabulary + property.Substring(colon + 1) : property;
    }
}
