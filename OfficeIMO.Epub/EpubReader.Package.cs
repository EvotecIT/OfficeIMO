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
        EpubDiagnosticCollector diagnostics,
        out IReadOnlyList<EpubRootfile> rootfiles) {
        rootfiles = ReadRootfiles(entryIndex, options, diagnostics);
        var attemptedPaths = new HashSet<string>(StringComparer.Ordinal);
        foreach (EpubRootfile rootfile in rootfiles) {
            if (!entryIndex.TryGetValue(rootfile.FullPath, out ZipArchiveEntry? opfEntry)) {
                diagnostics.Warning(
                    "epub.package.rootfile-missing",
                    $"Declared EPUB rootfile '{rootfile.FullPath}' was not found in archive.",
                    rootfile.FullPath);
                continue;
            }
            attemptedPaths.Add(rootfile.FullPath);
            EpubPackage? package = TryParsePackageEntry(opfEntry, rootfile.FullPath, options, diagnostics);
            if (package == null) continue;

            rootfile.IsSelected = true;
            return package;
        }

        foreach (string opfPath in entryIndex.Keys
            .Where(path => path.EndsWith(".opf", StringComparison.OrdinalIgnoreCase) && !attemptedPaths.Contains(path))
            .OrderBy(static path => path, StringComparer.Ordinal)) {
            EpubPackage? package = TryParsePackageEntry(entryIndex[opfPath], opfPath, options, diagnostics);
            if (package == null) continue;

            EpubRootfile? declared = rootfiles.FirstOrDefault(rootfile =>
                string.Equals(rootfile.FullPath, opfPath, StringComparison.Ordinal));
            if (declared != null) declared.IsSelected = true;
            diagnostics.Warning(
                "epub.container.rootfile-fallback",
                "No readable declared rootfile was selected. Falling back to the first readable discovered OPF.",
                opfPath);
            return package;
        }

        diagnostics.Warning(
            "epub.package.missing",
            "A readable EPUB OPF package document was not found.");
        return null;
    }

    private static EpubPackage? TryParsePackageEntry(
        ZipArchiveEntry opfEntry,
        string opfPath,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        if (opfEntry.Length > options.MaxPackageMetadataBytes) {
            diagnostics.Warning(
                "epub.package.metadata-size-limit",
                $"EPUB OPF package '{opfPath}' exceeds MaxPackageMetadataBytes ({options.MaxPackageMetadataBytes}).",
                opfPath);
            return null;
        }

        string opfContent = ReadEntryText(opfEntry, options.MaxPackageMetadataBytes);
        if (!TryParseXml(opfContent, out XDocument? opfDocument) || opfDocument == null) {
            diagnostics.Warning(
                "epub.package.invalid-xml",
                $"EPUB OPF package '{opfPath}' could not be parsed as XML.",
                opfPath);
            return null;
        }

        return ParseOpf(opfDocument, opfPath, options, diagnostics);
    }

    private static IReadOnlyList<EpubRootfile> ReadRootfiles(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        if (!entryIndex.TryGetValue("META-INF/container.xml", out var containerEntry)) {
            return Array.Empty<EpubRootfile>();
        }

        if (containerEntry.Length > options.MaxPackageMetadataBytes) {
            diagnostics.Warning(
                "epub.container.metadata-size-limit",
                $"EPUB container.xml exceeds MaxPackageMetadataBytes ({options.MaxPackageMetadataBytes}).",
                "META-INF/container.xml");
            return Array.Empty<EpubRootfile>();
        }

        var containerContent = ReadEntryText(containerEntry, options.MaxPackageMetadataBytes);
        if (!TryParseXml(containerContent, out var containerDocument) || containerDocument == null) {
            diagnostics.Warning(
                "epub.container.invalid-xml",
                "EPUB container.xml could not be parsed as XML.",
                "META-INF/container.xml");
            return Array.Empty<EpubRootfile>();
        }

        var results = new List<EpubRootfile>();
        var seenPaths = new HashSet<string>(StringComparer.Ordinal);
        foreach (var rootfile in containerDocument.Descendants().Where(e => IsName(e, "rootfile"))) {
            string declaredPath = GetAttribute(rootfile, "full-path");
            string candidate = RemoveFragmentAndQuery(declaredPath);
            if (!TryNormalizeArchiveEntryPath(candidate, out string fullPath)) {
                diagnostics.Warning(
                    "epub.container.rootfile-path-invalid",
                    $"Ignored rootfile declaration with invalid path '{declaredPath}'.",
                    "META-INF/container.xml");
                continue;
            }
            if (!seenPaths.Add(fullPath)) {
                diagnostics.Warning(
                    "epub.container.rootfile-duplicate",
                    $"Ignored duplicate rootfile declaration for '{fullPath}'.",
                    fullPath);
                continue;
            }

            results.Add(new EpubRootfile {
                FullPath = fullPath,
                MediaType = NullIfWhiteSpace(GetAttribute(rootfile, "media-type")),
                IsAvailable = entryIndex.ContainsKey(fullPath)
            });
        }

        if (results.Count == 0) {
            diagnostics.Warning(
                "epub.container.rootfile-missing",
                "EPUB container.xml did not define a valid rootfile path.",
                "META-INF/container.xml");
        } else if (results.Count > 1) {
            diagnostics.Warning(
                "epub.container.multiple-rootfiles",
                $"EPUB container declares {results.Count} rootfiles. The first readable package is selected.",
                "META-INF/container.xml");
        }
        return results.ToArray();
    }

    private static EpubPackage ParseOpf(
        XDocument opfDocument,
        string opfPath,
        EpubReadOptions options,
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
            ReadMetadataEntries(metadata, package, options, diagnostics, opfPath);
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

            bool isRemote = TryGetRemoteResourceUri(href, out string? remoteUri);
            string fullPath = isRemote ? remoteUri! : ResolveRelativePath(opfPath, href);
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
                Properties = GetAttribute(item, "properties"),
                IsRemote = isRemote,
                RemoteUri = remoteUri
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

        XElement? guide = opfDocument.Descendants().FirstOrDefault(element => IsName(element, "guide"));
        if (guide != null) {
            foreach (XElement reference in guide.Elements().Where(element => IsName(element, "reference"))) {
                if (package.Guide.Count >= options.MaxNavigationItems) {
                    diagnostics.Warning(
                        "epub.navigation.item-count-limit",
                        $"EPUB navigation was truncated at MaxNavigationItems ({options.MaxNavigationItems}).",
                        opfPath);
                    break;
                }
                string href = GetAttribute(reference, "href");
                if (!TryResolveNavigationTarget(opfPath, href, out string? target, out string? fragment, out bool isRemote)) {
                    diagnostics.Warning(
                        "epub.guide.invalid-target",
                        $"Ignored EPUB 2 guide reference with invalid href '{href}'.",
                        opfPath);
                    continue;
                }
                package.Guide.Add(new EpubNavigationItem {
                    Source = EpubNavigationSource.Epub2Guide,
                    Label = NullIfWhiteSpace(GetAttribute(reference, "title")) ?? GetAttribute(reference, "type"),
                    Href = href,
                    Target = target,
                    Fragment = fragment,
                    SemanticType = NullIfWhiteSpace(GetAttribute(reference, "type")),
                    IsRemote = isRemote
                });
            }
        }

        return package;
    }

    private static void ReadMetadataEntries(
        XElement metadata,
        EpubPackage package,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics,
        string opfPath) {
        foreach (XElement element in metadata.Elements()) {
            if (package.Metadata.Count >= options.MaxMetadataItems) {
                diagnostics.Warning(
                    "epub.metadata.count-limit",
                    $"EPUB metadata was truncated at MaxMetadataItems ({options.MaxMetadataItems}).",
                    opfPath);
                break;
            }

            string localName = element.Name.LocalName;
            EpubMetadataKind kind = string.Equals(element.Name.NamespaceName, "http://purl.org/dc/elements/1.1/", StringComparison.Ordinal)
                ? EpubMetadataKind.DublinCore
                : string.Equals(localName, "meta", StringComparison.OrdinalIgnoreCase)
                    ? EpubMetadataKind.Meta
                    : string.Equals(localName, "link", StringComparison.OrdinalIgnoreCase)
                        ? EpubMetadataKind.Link
                        : EpubMetadataKind.Other;
            string property = GetAttribute(element, "property");
            string legacyName = GetAttribute(element, "name");
            string href = GetAttribute(element, "href");
            string value = kind == EpubMetadataKind.Meta && property.Length == 0
                ? GetAttribute(element, "content")
                : kind == EpubMetadataKind.Link
                    ? href
                    : NormalizeWhitespace(element.Value);

            package.Metadata.Add(new EpubMetadataEntry {
                Kind = kind,
                Name = localName,
                NamespaceUri = element.Name.NamespaceName,
                Value = value,
                Id = NullIfWhiteSpace(GetAttribute(element, "id")),
                Property = NullIfWhiteSpace(property),
                Refines = NullIfWhiteSpace(GetAttribute(element, "refines")),
                Scheme = NullIfWhiteSpace(GetAttribute(element, "scheme")),
                Language = NullIfWhiteSpace(GetAttribute(element, "lang")),
                LegacyName = NullIfWhiteSpace(legacyName),
                Role = NullIfWhiteSpace(GetAttribute(element, "role")),
                FileAs = NullIfWhiteSpace(GetAttribute(element, "file-as")),
                Event = NullIfWhiteSpace(GetAttribute(element, "event")),
                Href = NullIfWhiteSpace(href),
                Rel = NullIfWhiteSpace(GetAttribute(element, "rel")),
                MediaType = NullIfWhiteSpace(GetAttribute(element, "media-type"))
            });
        }
    }

    private static bool TryGetRemoteResourceUri(string href, out string? remoteUri) {
        remoteUri = null;
        string candidate = href.Trim();
        if (!Uri.TryCreate(candidate, UriKind.Absolute, out Uri? uri)) return false;
        if (string.Equals(uri.Scheme, Uri.UriSchemeFile, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(uri.Scheme, "data", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }
        remoteUri = candidate;
        return true;
    }

}
