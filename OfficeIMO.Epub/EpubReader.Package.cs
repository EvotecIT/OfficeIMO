namespace OfficeIMO.Epub;

internal static partial class EpubReader {
    private static Dictionary<string, ZipArchiveEntry> BuildEntryIndex(
        ZipArchive archive,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        var map = new Dictionary<string, ZipArchiveEntry>(StringComparer.Ordinal);
        long totalUncompressedBytes = 0;
        int entryCount = 0;
        foreach (var entry in archive.Entries) {
            entryCount++;
            if (entryCount > options.MaxArchiveEntries) {
                throw CreateFatalReadException(
                    "epub.archive.entry-count-limit",
                    $"EPUB archive contains more than MaxArchiveEntries ({options.MaxArchiveEntries}) entries.");
            }

            try {
                totalUncompressedBytes = checked(totalUncompressedBytes + entry.Length);
            } catch (OverflowException exception) {
                throw CreateFatalReadException(
                    "epub.archive.total-size-limit",
                    "EPUB archive uncompressed size exceeds the supported range.",
                    null,
                    exception);
            }
            if (totalUncompressedBytes > options.MaxTotalUncompressedBytes) {
                throw CreateFatalReadException(
                    "epub.archive.total-size-limit",
                    $"EPUB archive uncompressed size exceeds MaxTotalUncompressedBytes ({options.MaxTotalUncompressedBytes}).");
            }

            if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue;

            if (!TryNormalizeArchiveEntryPath(entry.FullName, out string key)) {
                diagnostics.Warning(
                    "epub.archive.unsafe-path",
                    $"Ignored archive entry '{NormalizePath(entry.FullName)}' because its path is not safe.",
                    NormalizePath(entry.FullName));
                continue;
            }
            if (key.Length == 0) continue;
            if (map.ContainsKey(key)) {
                diagnostics.Warning(
                    "epub.archive.duplicate-path",
                    $"Ignored duplicate archive entry '{key}'.",
                    key);
                continue;
            }
            map[key] = entry;
        }

        return map;
    }

    private static EpubPackage? TryReadPackage(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        var opfPath = TryReadOpfPathFromContainer(entryIndex, options, diagnostics);
        if (opfPath == null) {
            opfPath = entryIndex.Keys
                .Where(k => k.EndsWith(".opf", StringComparison.OrdinalIgnoreCase))
                .OrderBy(k => k, StringComparer.Ordinal)
                .FirstOrDefault();

            if (opfPath != null) {
                diagnostics.Warning(
                    "epub.container.rootfile-fallback",
                    "EPUB container.xml rootfile was not found. Falling back to first discovered OPF.",
                    opfPath);
            }
        }

        if (opfPath == null) {
            diagnostics.Warning(
                "epub.package.missing",
                "EPUB OPF package document was not found.");
            return null;
        }

        if (!entryIndex.TryGetValue(opfPath, out var opfEntry)) {
            diagnostics.Warning(
                "epub.package.missing",
                $"EPUB OPF package '{opfPath}' was referenced but not found in archive.",
                opfPath);
            return null;
        }

        if (opfEntry.Length > options.MaxPackageMetadataBytes) {
            diagnostics.Warning(
                "epub.package.metadata-size-limit",
                $"EPUB OPF package '{opfPath}' exceeds MaxPackageMetadataBytes ({options.MaxPackageMetadataBytes}).",
                opfPath);
            return null;
        }

        var opfContent = ReadEntryText(opfEntry, options.MaxPackageMetadataBytes);
        if (!TryParseXml(opfContent, out var opfDocument) || opfDocument == null) {
            diagnostics.Warning(
                "epub.package.invalid-xml",
                $"EPUB OPF package '{opfPath}' could not be parsed as XML.",
                opfPath);
            return null;
        }

        return ParseOpf(opfDocument, opfPath, diagnostics);
    }

    private static string? TryReadOpfPathFromContainer(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        if (!entryIndex.TryGetValue("META-INF/container.xml", out var containerEntry)) {
            return null;
        }

        if (containerEntry.Length > options.MaxPackageMetadataBytes) {
            diagnostics.Warning(
                "epub.container.metadata-size-limit",
                $"EPUB container.xml exceeds MaxPackageMetadataBytes ({options.MaxPackageMetadataBytes}).",
                "META-INF/container.xml");
            return null;
        }

        var containerContent = ReadEntryText(containerEntry, options.MaxPackageMetadataBytes);
        if (!TryParseXml(containerContent, out var containerDocument) || containerDocument == null) {
            diagnostics.Warning(
                "epub.container.invalid-xml",
                "EPUB container.xml could not be parsed as XML.",
                "META-INF/container.xml");
            return null;
        }

        foreach (var rootfile in containerDocument.Descendants().Where(e => IsName(e, "rootfile"))) {
            var fullPath = GetAttribute(rootfile, "full-path");
            if (!string.IsNullOrWhiteSpace(fullPath)) {
                return NormalizePath(fullPath);
            }
        }

        diagnostics.Warning(
            "epub.container.rootfile-missing",
            "EPUB container.xml did not define a rootfile path.",
            "META-INF/container.xml");
        return null;
    }

    private static EpubPackage ParseOpf(
        XDocument opfDocument,
        string opfPath,
        EpubDiagnosticCollector diagnostics) {
        XElement? packageElement = opfDocument.Root;
        var package = new EpubPackage {
            OpfPath = opfPath,
            PackageVersion = packageElement == null ? null : NullIfWhiteSpace(GetAttribute(packageElement, "version")),
            UniqueIdentifierId = packageElement == null ? null : NullIfWhiteSpace(GetAttribute(packageElement, "unique-identifier"))
        };

        bool declaredUniqueIdentifierResolved = false;
        var metadata = opfDocument.Descendants().FirstOrDefault(e => IsName(e, "metadata"));
        if (metadata != null) {
            package.Title = TryGetFirstElementValue(metadata, "title");
            package.Creator = TryGetFirstElementValue(metadata, "creator");
            package.Language = TryGetFirstElementValue(metadata, "language");
            string? declaredIdentifier = null;
            if (!string.IsNullOrWhiteSpace(package.UniqueIdentifierId)) {
                XElement? declaredIdentifierElement = metadata.Elements().FirstOrDefault(element =>
                    IsName(element, "identifier") &&
                    string.Equals(GetAttribute(element, "id"), package.UniqueIdentifierId, StringComparison.Ordinal));
                if (declaredIdentifierElement != null) {
                    declaredIdentifier = NullIfWhiteSpace(NormalizeWhitespace(declaredIdentifierElement.Value));
                    declaredUniqueIdentifierResolved = declaredIdentifier != null;
                }
            }
            package.Identifier = declaredIdentifier ?? metadata.Elements()
                .Where(element => IsName(element, "identifier"))
                .Select(element => NullIfWhiteSpace(NormalizeWhitespace(element.Value)))
                .FirstOrDefault(identifier => identifier != null);

            package.RenditionLayout = ReadPackageRenditionLayout(metadata, diagnostics, opfPath);
        }

        if (string.IsNullOrWhiteSpace(package.PackageVersion)) {
            diagnostics.Warning(
                "epub.package.version-missing",
                "EPUB package does not declare a version.",
                opfPath);
        }
        if (!string.IsNullOrWhiteSpace(package.UniqueIdentifierId) && !declaredUniqueIdentifierResolved) {
            diagnostics.Warning(
                "epub.package.unique-identifier-missing",
                $"EPUB package unique-identifier '{package.UniqueIdentifierId}' does not reference a dc:identifier.",
                opfPath);
        }

        var manifestItems = opfDocument.Descendants().Where(e => IsName(e, "item"));
        foreach (var item in manifestItems) {
            var id = GetAttribute(item, "id");
            var href = GetAttribute(item, "href");
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(href)) continue;

            string fullPath = ResolveRelativePath(opfPath, href);
            if (fullPath.Length == 0) {
                diagnostics.Warning(
                    "epub.manifest.invalid-path",
                    $"Ignored manifest item '{id}' because href '{href}' does not resolve to a safe archive path.",
                    opfPath);
                continue;
            }
            var model = new ManifestItem {
                Id = id,
                Href = href,
                FullPath = fullPath,
                MediaType = GetAttribute(item, "media-type"),
                Properties = GetAttribute(item, "properties")
            };
            if (package.Manifest.ContainsKey(id)) {
                diagnostics.Warning(
                    "epub.manifest.duplicate-id",
                    $"EPUB manifest contains duplicate id '{id}'. The last declaration is used.",
                    opfPath);
            }
            package.Manifest[id] = model;

            if (ContainsSpaceSeparatedToken(model.Properties, "nav")) {
                package.NavDocumentPath = model.FullPath;
            }
            if (!string.IsNullOrWhiteSpace(model.MediaType) &&
                model.MediaType.IndexOf("ncx", StringComparison.OrdinalIgnoreCase) >= 0 &&
                string.IsNullOrWhiteSpace(package.NcxPath)) {
                package.NcxPath = model.FullPath;
            }
        }

        var spine = opfDocument.Descendants().FirstOrDefault(e => IsName(e, "spine"));
        if (spine != null) {
            var tocId = GetAttribute(spine, "toc");
            if (!string.IsNullOrWhiteSpace(tocId) &&
                package.Manifest.TryGetValue(tocId, out var tocManifest) &&
                string.IsNullOrWhiteSpace(package.NcxPath)) {
                package.NcxPath = tocManifest.FullPath;
            }

            int index = 0;
            foreach (var itemRef in spine.Elements().Where(e => IsName(e, "itemref"))) {
                index++;
                var idRef = GetAttribute(itemRef, "idref");
                if (string.IsNullOrWhiteSpace(idRef)) continue;

                var linear = GetAttribute(itemRef, "linear");
                var isLinear = !string.Equals(linear, "no", StringComparison.OrdinalIgnoreCase);
                string properties = GetAttribute(itemRef, "properties");

                package.Spine.Add(new SpineItem {
                    IdRef = idRef,
                    SpineIndex = index,
                    IsLinear = isLinear,
                    Properties = properties,
                    RenditionLayout = ResolveRenditionLayout(package.RenditionLayout, properties)
                });
            }
        }

        return package;
    }

}
