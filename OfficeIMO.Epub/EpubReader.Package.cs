namespace OfficeIMO.Epub;

using System.Threading;

internal static partial class EpubReader {
    private static Dictionary<string, ZipArchiveEntry> BuildEntryIndex(
        ZipArchive archive,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics,
        CancellationToken cancellationToken) {
        var map = new Dictionary<string, ZipArchiveEntry>(StringComparer.Ordinal);
        long totalUncompressedBytes = 0;
        int entryCount = 0;
        foreach (var entry in archive.Entries) {
            cancellationToken.ThrowIfCancellationRequested();
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
        out IReadOnlyList<EpubRootfile> rootfiles,
        CancellationToken cancellationToken) {
        rootfiles = ReadRootfiles(entryIndex, options, diagnostics, cancellationToken);
        var attemptedPaths = new HashSet<string>(StringComparer.Ordinal);
        foreach (EpubRootfile rootfile in rootfiles) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!entryIndex.TryGetValue(rootfile.FullPath, out ZipArchiveEntry? opfEntry)) {
                diagnostics.Warning(
                    "epub.package.rootfile-missing",
                    $"Declared EPUB rootfile '{rootfile.FullPath}' was not found in archive.",
                    rootfile.FullPath);
                continue;
            }
            attemptedPaths.Add(rootfile.FullPath);
            EpubPackage? package = TryParsePackageEntry(opfEntry, rootfile.FullPath, options, diagnostics, cancellationToken);
            if (package == null) continue;

            rootfile.IsSelected = true;
            return package;
        }

        foreach (string opfPath in entryIndex.Keys
            .Where(path => path.EndsWith(".opf", StringComparison.OrdinalIgnoreCase) && !attemptedPaths.Contains(path))
            .OrderBy(static path => path, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            EpubPackage? package = TryParsePackageEntry(entryIndex[opfPath], opfPath, options, diagnostics, cancellationToken);
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
        EpubDiagnosticCollector diagnostics,
        CancellationToken cancellationToken) {
        if (opfEntry.Length > options.MaxPackageMetadataBytes) {
            diagnostics.Warning(
                "epub.package.metadata-size-limit",
                $"EPUB OPF package '{opfPath}' exceeds MaxPackageMetadataBytes ({options.MaxPackageMetadataBytes}).",
                opfPath);
            return null;
        }

        if (!TryParseEntryXml(opfEntry, options.MaxPackageMetadataBytes, cancellationToken, out XDocument? opfDocument)
            || opfDocument == null) {
            diagnostics.Warning(
                "epub.package.invalid-xml",
                $"EPUB OPF package '{opfPath}' could not be parsed as XML.",
                opfPath);
            return null;
        }

        if (opfDocument.Root == null || !IsOpfName(opfDocument.Root, "package")) {
            diagnostics.Warning(
                "epub.package.namespace-invalid",
                "EPUB OPF package root must be a package element in the OPF namespace.",
                opfPath);
            return null;
        }

        return ParseOpf(opfDocument, opfPath, options, diagnostics, cancellationToken);
    }

    private static IReadOnlyList<EpubRootfile> ReadRootfiles(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics,
        CancellationToken cancellationToken) {
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

        if (!TryParseEntryXml(containerEntry, options.MaxPackageMetadataBytes, cancellationToken, out XDocument? containerDocument)
            || containerDocument == null) {
            diagnostics.Warning(
                "epub.container.invalid-xml",
                "EPUB container.xml could not be parsed as XML.",
                "META-INF/container.xml");
            return Array.Empty<EpubRootfile>();
        }

        var results = new List<EpubRootfile>();
        var seenPaths = new HashSet<string>(StringComparer.Ordinal);
        foreach (var rootfile in containerDocument.Descendants().Where(e => IsContainerName(e, "rootfile"))) {
            cancellationToken.ThrowIfCancellationRequested();
            string declaredPath = GetUnqualifiedAttribute(rootfile, "full-path");
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
                MediaType = NullIfWhiteSpace(GetUnqualifiedAttribute(rootfile, "media-type")),
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
        EpubDiagnosticCollector diagnostics,
        CancellationToken cancellationToken) {
        XElement? packageElement = opfDocument.Root;
        var package = new EpubPackage {
            OpfPath = opfPath,
            PackageVersion = packageElement == null ? null : NullIfWhiteSpace(GetUnqualifiedAttribute(packageElement, "version")),
            UniqueIdentifierId = packageElement == null ? null : NullIfWhiteSpace(GetUnqualifiedAttribute(packageElement, "unique-identifier"))
        };

        bool declaredUniqueIdentifierResolved = false;
        XElement? metadata = packageElement?.Elements().FirstOrDefault(e => IsOpfName(e, "metadata"));
        if (metadata != null) {
            ReadMetadataEntries(metadata, package, options, diagnostics, opfPath, cancellationToken);
            package.Title = TryGetFirstDublinCoreValue(metadata, "title");
            package.Creator = TryGetFirstDublinCoreValue(metadata, "creator");
            package.Language = TryGetFirstDublinCoreValue(metadata, "language");
            string? declaredIdentifier = null;
            if (!string.IsNullOrWhiteSpace(package.UniqueIdentifierId)) {
                XElement? declaredIdentifierElement = metadata.Elements().FirstOrDefault(element =>
                    IsDublinCoreName(element, "identifier") &&
                    string.Equals(GetUnqualifiedAttribute(element, "id"), package.UniqueIdentifierId, StringComparison.Ordinal));
                if (declaredIdentifierElement != null) {
                    package.ObfuscationIdentifier = declaredIdentifierElement.Value;
                    declaredIdentifier = NullIfWhiteSpace(NormalizeWhitespace(declaredIdentifierElement.Value));
                    declaredUniqueIdentifierResolved = declaredIdentifier != null;
                }
            }
            package.Identifier = declaredIdentifier ?? metadata.Elements()
                .Where(element => IsDublinCoreName(element, "identifier"))
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

        var manifestTargets = new HashSet<string>(StringComparer.Ordinal);
        XElement? manifest = packageElement?.Elements().FirstOrDefault(e => IsOpfName(e, "manifest"));
        IEnumerable<XElement> manifestItems = manifest?.Elements().Where(e => IsOpfName(e, "item"))
            ?? Enumerable.Empty<XElement>();
        foreach (var item in manifestItems) {
            cancellationToken.ThrowIfCancellationRequested();
            var id = GetUnqualifiedAttribute(item, "id");
            var href = GetUnqualifiedAttribute(item, "href");
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(href)) continue;

            EpubReference resolvedReference = EpubReference.Resolve(opfPath, href);
            bool isRemote = resolvedReference.Kind == EpubReferenceKind.External;
            string? remoteUri = isRemote ? resolvedReference.ResolvedValue : null;
            string fullPath = isRemote
                ? remoteUri ?? string.Empty
                : resolvedReference.Kind == EpubReferenceKind.Container
                    ? resolvedReference.ContainerPath ?? string.Empty
                    : string.Empty;
            if (fullPath.Length == 0) {
                diagnostics.Warning(
                    "epub.manifest.invalid-path",
                    $"Ignored manifest item '{id}' because href '{href}' does not resolve to a safe archive path.",
                    opfPath);
                continue;
            }
            if (!resolvedReference.IsConforming) {
                diagnostics.Warning(
                    "epub.manifest.reference-non-conforming",
                    $"Manifest item '{id}' uses non-conforming root-relative href '{href}'. The safe container target is retained.",
                    fullPath);
            }
            var model = new ManifestItem {
                Id = id,
                Href = href,
                FullPath = fullPath,
                MediaType = GetUnqualifiedAttribute(item, "media-type"),
                Properties = GetUnqualifiedAttribute(item, "properties"),
                IsRemote = isRemote,
                RemoteUri = remoteUri
            };
            if (package.Manifest.ContainsKey(id)) {
                diagnostics.Warning(
                    "epub.manifest.duplicate-id",
                    $"EPUB manifest contains duplicate id '{id}'. The last declaration is used.",
                    opfPath);
            }
            if (!isRemote && !manifestTargets.Add(fullPath)) {
                diagnostics.Warning(
                    "epub.manifest.duplicate-target",
                    $"EPUB manifest resolves more than one item to archive path '{fullPath}'.",
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

        XElement? spine = packageElement?.Elements().FirstOrDefault(e => IsOpfName(e, "spine"));
        if (spine != null) {
            var tocId = GetUnqualifiedAttribute(spine, "toc");
            if (!string.IsNullOrWhiteSpace(tocId) &&
                package.Manifest.TryGetValue(tocId, out var tocManifest) &&
                string.IsNullOrWhiteSpace(package.NcxPath)) {
                package.NcxPath = tocManifest.FullPath;
            }

            int index = 0;
            foreach (var itemRef in spine.Elements().Where(e => IsOpfName(e, "itemref"))) {
                cancellationToken.ThrowIfCancellationRequested();
                index++;
                var idRef = GetUnqualifiedAttribute(itemRef, "idref");
                if (string.IsNullOrWhiteSpace(idRef)) continue;

                var linear = GetUnqualifiedAttribute(itemRef, "linear");
                var isLinear = !string.Equals(linear, "no", StringComparison.OrdinalIgnoreCase);
                string properties = GetUnqualifiedAttribute(itemRef, "properties");

                package.Spine.Add(new SpineItem {
                    IdRef = idRef,
                    SpineIndex = index,
                    IsLinear = isLinear,
                    Properties = properties,
                    RenditionLayout = ResolveRenditionLayout(package.RenditionLayout, properties)
                });
            }
        }

        XElement? guide = opfDocument.Descendants().FirstOrDefault(element => IsOpfName(element, "guide"));
        if (guide != null) {
            foreach (XElement reference in guide.Elements().Where(element => IsOpfName(element, "reference"))) {
                cancellationToken.ThrowIfCancellationRequested();
                if (package.Guide.Count >= options.MaxNavigationItems) {
                    diagnostics.Warning(
                        "epub.navigation.item-count-limit",
                        $"EPUB navigation was truncated at MaxNavigationItems ({options.MaxNavigationItems}).",
                        opfPath);
                    break;
                }
                string href = GetUnqualifiedAttribute(reference, "href");
                if (!TryResolveNavigationTarget(opfPath, null, href, out string? target, out string? fragment, out bool isRemote)) {
                    diagnostics.Warning(
                        "epub.guide.invalid-target",
                        $"Ignored EPUB 2 guide reference with invalid href '{href}'.",
                        opfPath);
                    continue;
                }
                package.Guide.Add(new EpubNavigationItem {
                    Source = EpubNavigationSource.Epub2Guide,
                    Label = NullIfWhiteSpace(GetUnqualifiedAttribute(reference, "title")) ?? GetUnqualifiedAttribute(reference, "type"),
                    Href = href,
                    Target = target,
                    Fragment = fragment,
                    SemanticType = NullIfWhiteSpace(GetUnqualifiedAttribute(reference, "type")),
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
        string opfPath,
        CancellationToken cancellationToken) {
        foreach (XElement element in metadata.Elements()) {
            cancellationToken.ThrowIfCancellationRequested();
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
            string property = GetUnqualifiedAttribute(element, "property");
            string legacyName = GetUnqualifiedAttribute(element, "name");
            string href = GetUnqualifiedAttribute(element, "href");
            string value = kind == EpubMetadataKind.Meta && property.Length == 0
                ? GetUnqualifiedAttribute(element, "content")
                : kind == EpubMetadataKind.Link
                    ? href
                    : NormalizeWhitespace(element.Value);

            package.Metadata.Add(new EpubMetadataEntry {
                Kind = kind,
                Name = localName,
                NamespaceUri = element.Name.NamespaceName,
                Value = value,
                Id = NullIfWhiteSpace(GetUnqualifiedAttribute(element, "id")),
                Property = NullIfWhiteSpace(property),
                Refines = NullIfWhiteSpace(GetUnqualifiedAttribute(element, "refines")),
                Scheme = NullIfWhiteSpace(GetOpfMetadataAttribute(element, "scheme")),
                Language = NullIfWhiteSpace(GetXmlLanguage(element)),
                LegacyName = NullIfWhiteSpace(legacyName),
                Role = NullIfWhiteSpace(GetOpfMetadataAttribute(element, "role")),
                FileAs = NullIfWhiteSpace(GetOpfMetadataAttribute(element, "file-as")),
                Event = NullIfWhiteSpace(GetOpfMetadataAttribute(element, "event")),
                Href = NullIfWhiteSpace(href),
                Rel = NullIfWhiteSpace(GetUnqualifiedAttribute(element, "rel")),
                MediaType = NullIfWhiteSpace(GetUnqualifiedAttribute(element, "media-type"))
            });
        }
    }

}
