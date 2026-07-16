namespace OfficeIMO.Epub;

internal static partial class EpubReader {
    private static Dictionary<string, string> BuildNavigationTitleMap(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        EpubPackage? package,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        if (package == null) return map;

        if (!string.IsNullOrWhiteSpace(package.NavDocumentPath)) {
            TryReadNavigationDocument(entryIndex, package.NavDocumentPath!, map, options, diagnostics);
        }
        if (!string.IsNullOrWhiteSpace(package.NcxPath)) {
            TryReadNcxDocument(entryIndex, package.NcxPath!, map, options, diagnostics);
        }

        return map;
    }

    private static void TryReadNavigationDocument(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        string navPath,
        Dictionary<string, string> map,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        if (!entryIndex.TryGetValue(navPath, out var navEntry)) {
            diagnostics.Warning(
                "epub.navigation.missing",
                $"EPUB navigation document '{navPath}' was not found in archive.",
                navPath);
            return;
        }

        if (navEntry.Length > options.MaxPackageMetadataBytes) {
            diagnostics.Warning(
                "epub.navigation.metadata-size-limit",
                $"EPUB navigation document '{navPath}' exceeds MaxPackageMetadataBytes ({options.MaxPackageMetadataBytes}).",
                navPath);
            return;
        }

        var navContent = ReadEntryText(navEntry, options.MaxPackageMetadataBytes);
        if (!TryParseXml(navContent, out var navDocument) || navDocument == null) {
            diagnostics.Warning(
                "epub.navigation.invalid-xml",
                $"EPUB navigation document '{navPath}' could not be parsed.",
                navPath);
            return;
        }

        foreach (var anchor in navDocument.Descendants().Where(e => IsName(e, "a"))) {
            var href = GetAttribute(anchor, "href");
            if (string.IsNullOrWhiteSpace(href)) continue;

            var targetPath = ResolveRelativePath(navPath, href);
            if (targetPath.Length == 0) continue;

            var title = NormalizeWhitespace(anchor.Value);
            if (title.Length == 0) continue;

            if (!map.ContainsKey(targetPath)) {
                map[targetPath] = title;
            }
        }
    }

    private static void TryReadNcxDocument(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        string ncxPath,
        Dictionary<string, string> map,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        if (!entryIndex.TryGetValue(ncxPath, out var ncxEntry)) {
            diagnostics.Warning(
                "epub.ncx.missing",
                $"EPUB NCX document '{ncxPath}' was not found in archive.",
                ncxPath);
            return;
        }

        if (ncxEntry.Length > options.MaxPackageMetadataBytes) {
            diagnostics.Warning(
                "epub.ncx.metadata-size-limit",
                $"EPUB NCX document '{ncxPath}' exceeds MaxPackageMetadataBytes ({options.MaxPackageMetadataBytes}).",
                ncxPath);
            return;
        }

        var ncxContent = ReadEntryText(ncxEntry, options.MaxPackageMetadataBytes);
        if (!TryParseXml(ncxContent, out var ncxDocument) || ncxDocument == null) {
            diagnostics.Warning(
                "epub.ncx.invalid-xml",
                $"EPUB NCX document '{ncxPath}' could not be parsed.",
                ncxPath);
            return;
        }

        foreach (var navPoint in ncxDocument.Descendants().Where(e => IsName(e, "navPoint"))) {
            var content = navPoint.Descendants().FirstOrDefault(e => IsName(e, "content"));
            if (content == null) continue;

            var src = GetAttribute(content, "src");
            if (string.IsNullOrWhiteSpace(src)) continue;

            var targetPath = ResolveRelativePath(ncxPath, src);
            if (targetPath.Length == 0) continue;

            var textElement = navPoint.Descendants().FirstOrDefault(e => IsName(e, "text"));
            if (textElement == null) continue;

            var title = NormalizeWhitespace(textElement.Value);
            if (title.Length == 0) continue;

            if (!map.ContainsKey(targetPath)) {
                map[targetPath] = title;
            }
        }
    }
}
