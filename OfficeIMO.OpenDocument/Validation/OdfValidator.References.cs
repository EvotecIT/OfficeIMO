namespace OfficeIMO.OpenDocument;

internal static partial class OdfValidator {
    private static void ValidatePackageReferences(OdfPackage package, List<OdfDiagnostic> diagnostics) {
        foreach (string partPath in new[] { "content.xml", "styles.xml" }) {
            if (!package.ContainsEntry(partPath)) continue;
            XDocument document = package.GetXml(partPath);
            foreach (XElement owner in document.Descendants().Where(IsPackageReferenceOwner)) {
                string? href = (string?)owner.Attribute(OdfNamespaces.XLink + "href");
                if (string.IsNullOrWhiteSpace(href) || IsExternalOrFragment(href!)) continue;
                string normalized = NormalizePackageHref(href!);
                if (normalized.Length == 0) continue;
                bool exists = package.ContainsEntry(normalized) || package.Entries.Any(entry =>
                    entry.Name.StartsWith(normalized.TrimEnd('/') + "/", StringComparison.Ordinal));
                if (!exists) {
                    diagnostics.Add(new OdfDiagnostic("ODF300", OdfDiagnosticSeverity.Error,
                        $"{owner.Name.LocalName} references missing package entry '{href}'.", partPath));
                }
            }
        }
    }

    private static bool IsPackageReferenceOwner(XElement element) =>
        element.Name == OdfNamespaces.Draw + "image" || element.Name == OdfNamespaces.Draw + "object" ||
        element.Name == OdfNamespaces.Draw + "object-ole" || element.Name == OdfNamespaces.Draw + "plugin" ||
        element.Name == OdfNamespaces.Style + "background-image";

    private static bool IsExternalOrFragment(string href) => href.StartsWith("#", StringComparison.Ordinal) ||
        href.StartsWith("//", StringComparison.Ordinal) || Uri.TryCreate(href, UriKind.Absolute, out _);

    private static string NormalizePackageHref(string href) {
        string value = href;
        int fragment = value.IndexOf('#');
        if (fragment >= 0) value = value.Substring(0, fragment);
        int query = value.IndexOf('?');
        if (query >= 0) value = value.Substring(0, query);
        while (value.StartsWith("./", StringComparison.Ordinal)) value = value.Substring(2);
        try { return Uri.UnescapeDataString(value); } catch (UriFormatException) { return value; }
    }
}
