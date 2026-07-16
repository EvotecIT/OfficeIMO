namespace OfficeIMO.Epub;

using OfficeIMO.Drawing.Internal;

internal static partial class EpubReader {
    private static List<ChapterCandidate> BuildChapterCandidates(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        EpubPackage? package,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics) {
        var candidates = new List<ChapterCandidate>();
        var seenPaths = new HashSet<string>(StringComparer.Ordinal);

        if (package != null && options.PreferSpineOrder && package.Spine.Count > 0) {
            foreach (var spineItem in package.Spine.OrderBy(s => s.SpineIndex)) {
                if (!options.IncludeNonLinearSpineItems && !spineItem.IsLinear) {
                    continue;
                }

                if (!package.Manifest.TryGetValue(spineItem.IdRef, out var manifestItem)) {
                    diagnostics.Warning(
                        "epub.spine.manifest-id-missing",
                        $"EPUB spine idref '{spineItem.IdRef}' does not exist in manifest.",
                        package.OpfPath);
                    continue;
                }

                if (!IsChapterManifestItem(manifestItem)) {
                    continue;
                }

                if (!entryIndex.TryGetValue(manifestItem.FullPath, out var chapterEntry)) {
                    diagnostics.Warning(
                        "epub.spine.resource-missing",
                        $"EPUB manifest item '{manifestItem.FullPath}' referenced by spine was not found in archive.",
                        manifestItem.FullPath);
                    continue;
                }

                var chapterPath = NormalizePath(chapterEntry.FullName);
                if (seenPaths.Contains(chapterPath)) {
                    continue;
                }

                seenPaths.Add(chapterPath);
                candidates.Add(new ChapterCandidate {
                    Entry = chapterEntry,
                    ManifestId = manifestItem.Id,
                    MediaType = manifestItem.MediaType,
                    SpineIndex = spineItem.SpineIndex,
                    IsLinear = spineItem.IsLinear,
                    RenditionLayout = spineItem.RenditionLayout
                });
            }
        }

        var shouldFallbackScan = candidates.Count == 0 && options.FallbackToHtmlScan;
        if (!options.PreferSpineOrder) {
            shouldFallbackScan = true;
        }

        if (shouldFallbackScan) {
            IEnumerable<ZipArchiveEntry> scanEntries = entryIndex.Values.Where(e => IsChapterEntry(e.FullName));
            if (options.DeterministicOrder) {
                scanEntries = scanEntries.OrderBy(e => e.FullName, StringComparer.Ordinal);
            }

            var manifestByPath = BuildManifestByPath(package);
            foreach (var entry in scanEntries) {
                var chapterPath = NormalizePath(entry.FullName);
                if (seenPaths.Contains(chapterPath)) continue;

                manifestByPath.TryGetValue(chapterPath, out var manifestItem);
                seenPaths.Add(chapterPath);
                candidates.Add(new ChapterCandidate {
                    Entry = entry,
                    ManifestId = manifestItem?.Id,
                    MediaType = manifestItem?.MediaType,
                    SpineIndex = null,
                    IsLinear = null,
                    RenditionLayout = package?.RenditionLayout
                });
            }
        }

        return candidates;
    }

    private static Dictionary<string, ManifestItem> BuildManifestByPath(EpubPackage? package) {
        var map = new Dictionary<string, ManifestItem>(StringComparer.Ordinal);
        if (package == null) return map;

        foreach (var item in package.Manifest.Values) {
            if (!map.ContainsKey(item.FullPath)) {
                map[item.FullPath] = item;
            }
        }

        return map;
    }

    private static bool IsChapterManifestItem(ManifestItem item) {
        if (!string.IsNullOrWhiteSpace(item.MediaType) &&
            (item.MediaType.IndexOf("xhtml", StringComparison.OrdinalIgnoreCase) >= 0 ||
             item.MediaType.IndexOf("html", StringComparison.OrdinalIgnoreCase) >= 0)) {
            return true;
        }

        return IsChapterEntry(item.FullPath);
    }

    private static bool IsChapterEntry(string? fullName) {
        if (string.IsNullOrWhiteSpace(fullName)) return false;
        var normalized = NormalizePath(fullName!);
        if (normalized.EndsWith("/", StringComparison.Ordinal)) return false;

        var ext = Path.GetExtension(normalized).ToLowerInvariant();
        return ext == ".xhtml" || ext == ".html" || ext == ".htm";
    }

    private static string ReadEntryText(ZipArchiveEntry entry, long? maxBytes) {
        byte[] data;
        using (Stream entryStream = entry.Open()) {
            data = OfficeStreamReader.ReadAllBytes(entryStream, maxBytes);
        }
        using var memory = new MemoryStream(data, writable: false);
        using var reader = new StreamReader(memory, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 16 * 1024, leaveOpen: false);
        return reader.ReadToEnd();
    }

    private static byte[] ReadEntryBytes(ZipArchiveEntry entry, long maxBytes) {
        using Stream entryStream = entry.Open();
        return OfficeStreamReader.ReadAllBytes(entryStream, maxBytes);
    }

    private static bool TryParseXml(string content, out XDocument? document) {
        document = null;
        if (string.IsNullOrWhiteSpace(content)) return false;

        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Ignore,
                XmlResolver = null
            };

            using var stringReader = new StringReader(content);
            using var xmlReader = XmlReader.Create(stringReader, settings);
            document = XDocument.Load(xmlReader, LoadOptions.PreserveWhitespace);
            return true;
        } catch {
            return false;
        }
    }

    private static string ExtractVisibleText(XDocument chapterDocument) {
        var root = chapterDocument.Root;
        if (root == null) return string.Empty;

        var body = root.Descendants().FirstOrDefault(e => IsName(e, "body"));
        var scope = (XContainer?)body ?? root;

        var sb = new StringBuilder();
        foreach (var textNode in scope.DescendantNodes().OfType<XText>()) {
            if (textNode.Parent != null && (IsName(textNode.Parent, "script") || IsName(textNode.Parent, "style"))) {
                continue;
            }

            var value = textNode.Value;
            if (string.IsNullOrWhiteSpace(value)) continue;

            sb.Append(value);
            sb.Append(' ');
        }

        return NormalizeWhitespace(sb.ToString());
    }

    private static string? ResolveChapterTitle(XDocument chapterDocument, Dictionary<string, string> navTitleMap, string chapterPath) {
        if (navTitleMap.TryGetValue(chapterPath, out var navTitle) && !string.IsNullOrWhiteSpace(navTitle)) {
            return navTitle;
        }

        var title = chapterDocument.Descendants()
            .FirstOrDefault(e => IsName(e, "title"));
        if (title != null) {
            var normalized = NormalizeWhitespace(title.Value);
            if (normalized.Length > 0) return normalized;
        }

        var heading = chapterDocument.Descendants()
            .FirstOrDefault(e => IsName(e, "h1") || IsName(e, "h2"));
        if (heading != null) {
            var normalized = NormalizeWhitespace(heading.Value);
            if (normalized.Length > 0) return normalized;
        }

        return null;
    }

    private static string? ResolveDocumentTitle(EpubPackage? package, IReadOnlyList<EpubChapter> chapters) {
        if (!string.IsNullOrWhiteSpace(package?.Title)) {
            return package!.Title;
        }

        foreach (var chapter in chapters) {
            if (!string.IsNullOrWhiteSpace(chapter.Title)) {
                return chapter.Title;
            }
        }

        return null;
    }

}
