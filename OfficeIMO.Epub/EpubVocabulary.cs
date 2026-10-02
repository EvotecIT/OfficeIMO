namespace OfficeIMO.Epub;

/// <summary>Shared package vocabulary expansion for extraction and authoring.</summary>
internal static class EpubVocabulary {
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
        if (colon < 0) return "http://idpf.org/epub/vocab/package/#" + property;
        var prefixes = ReadPrefixes((string?)package.Attribute("prefix") ?? string.Empty);
        return prefixes.TryGetValue(property.Substring(0, colon), out string? vocabulary) ? vocabulary + property.Substring(colon + 1) : property;
    }
}
