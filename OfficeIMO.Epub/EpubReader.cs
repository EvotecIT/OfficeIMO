namespace OfficeIMO.Epub;

/// <summary>
/// Standards-based EPUB extractor using container/OPF/spine/navigation metadata.
/// </summary>
public static class EpubReader {
    /// <summary>
    /// Reads an EPUB document from disk.
    /// </summary>
    public static EpubDocument Read(string epubPath, EpubReadOptions? options = null) {
        if (epubPath == null) throw new ArgumentNullException(nameof(epubPath));
        if (epubPath.Length == 0) throw new ArgumentException("EPUB path cannot be empty.", nameof(epubPath));
        if (!File.Exists(epubPath)) throw new FileNotFoundException($"EPUB file '{epubPath}' doesn't exist.", epubPath);

        using var fs = new FileStream(epubPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Read(fs, options);
    }

    /// <summary>
    /// Reads an EPUB document from a stream.
    /// </summary>
    public static EpubDocument Read(Stream epubStream, EpubReadOptions? options = null) {
        if (epubStream == null) throw new ArgumentNullException(nameof(epubStream));
        if (!epubStream.CanRead) throw new IOException("EPUB stream must be readable.");

        var effective = Normalize(options);
        var warnings = new List<string>();

        using var archive = new ZipArchive(epubStream, ZipArchiveMode.Read, leaveOpen: true);
        var entryIndex = BuildEntryIndex(archive);
        var package = TryReadPackage(entryIndex, warnings);
        var navTitleMap = BuildNavigationTitleMap(entryIndex, package, warnings);
        var candidates = BuildChapterCandidates(archive, entryIndex, package, effective, warnings);

        var chapters = new List<EpubChapter>();
        int emitted = 0;

        foreach (var candidate in candidates) {
            if (emitted >= effective.MaxChapters) break;
            if (effective.MaxChapterBytes.HasValue && candidate.Entry.Length > effective.MaxChapterBytes.Value) {
                warnings.Add($"Skipped chapter '{NormalizePath(candidate.Entry.FullName)}' because size {candidate.Entry.Length} exceeds MaxChapterBytes ({effective.MaxChapterBytes.Value}).");
                continue;
            }

            var markup = ReadEntryText(candidate.Entry);
            if (!TryParseXml(markup, out var chapterDocument) || chapterDocument == null) {
                warnings.Add($"Skipped chapter '{NormalizePath(candidate.Entry.FullName)}' because chapter markup is not valid XML/XHTML.");
                continue;
            }

            var text = ExtractVisibleText(chapterDocument);
            if (text.Length == 0) {
                continue;
            }

            emitted++;
            var normalizedPath = NormalizePath(candidate.Entry.FullName);
            var title = ResolveChapterTitle(chapterDocument, navTitleMap, normalizedPath);

            chapters.Add(new EpubChapter {
                Order = emitted,
                Path = normalizedPath,
                ManifestId = candidate.ManifestId,
                MediaType = candidate.MediaType,
                SpineIndex = candidate.SpineIndex,
                IsLinear = candidate.IsLinear,
                Title = title,
                Text = text,
                Html = effective.IncludeRawHtml ? markup : null
            });
        }

        return new EpubDocument {
            Title = ResolveDocumentTitle(package, chapters),
            Identifier = package?.Identifier,
            Language = package?.Language,
            Creator = package?.Creator,
            OpfPath = package?.OpfPath,
            Chapters = chapters,
            Warnings = warnings
        };
    }

    private static Dictionary<string, ZipArchiveEntry> BuildEntryIndex(ZipArchive archive) {
        var map = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries) {
            var key = NormalizePath(entry.FullName);
            if (key.Length == 0) continue;
            if (!map.ContainsKey(key)) {
                map[key] = entry;
            }
        }

        return map;
    }

    private static EpubPackage? TryReadPackage(Dictionary<string, ZipArchiveEntry> entryIndex, List<string> warnings) {
        var opfPath = TryReadOpfPathFromContainer(entryIndex, warnings);
        if (opfPath == null) {
            opfPath = entryIndex.Keys
                .Where(k => k.EndsWith(".opf", StringComparison.OrdinalIgnoreCase))
                .OrderBy(k => k, StringComparer.Ordinal)
                .FirstOrDefault();

            if (opfPath != null) {
                warnings.Add("EPUB container.xml rootfile was not found. Falling back to first discovered OPF.");
            }
        }

        if (opfPath == null) {
            warnings.Add("EPUB OPF package document was not found.");
            return null;
        }

        if (!entryIndex.TryGetValue(opfPath, out var opfEntry)) {
            warnings.Add($"EPUB OPF package '{opfPath}' was referenced but not found in archive.");
            return null;
        }

        var opfContent = ReadEntryText(opfEntry);
        if (!TryParseXml(opfContent, out var opfDocument) || opfDocument == null) {
            warnings.Add($"EPUB OPF package '{opfPath}' could not be parsed as XML.");
            return null;
        }

        return ParseOpf(opfDocument, opfPath);
    }

    private static string? TryReadOpfPathFromContainer(Dictionary<string, ZipArchiveEntry> entryIndex, List<string> warnings) {
        if (!entryIndex.TryGetValue("META-INF/container.xml", out var containerEntry)) {
            return null;
        }

        var containerContent = ReadEntryText(containerEntry);
        if (!TryParseXml(containerContent, out var containerDocument) || containerDocument == null) {
            warnings.Add("EPUB container.xml could not be parsed as XML.");
            return null;
        }

        foreach (var rootfile in containerDocument.Descendants().Where(e => IsName(e, "rootfile"))) {
            var fullPath = GetAttribute(rootfile, "full-path");
            if (!string.IsNullOrWhiteSpace(fullPath)) {
                return NormalizePath(fullPath);
            }
        }

        warnings.Add("EPUB container.xml did not define a rootfile path.");
        return null;
    }

    private static EpubPackage ParseOpf(XDocument opfDocument, string opfPath) {
        var package = new EpubPackage {
            OpfPath = opfPath
        };

        var metadata = opfDocument.Descendants().FirstOrDefault(e => IsName(e, "metadata"));
        if (metadata != null) {
            package.Title = TryGetFirstElementValue(metadata, "title");
            package.Creator = TryGetFirstElementValue(metadata, "creator");
            package.Language = TryGetFirstElementValue(metadata, "language");
            package.Identifier = TryGetFirstElementValue(metadata, "identifier");
        }

        var manifestItems = opfDocument.Descendants().Where(e => IsName(e, "item"));
        foreach (var item in manifestItems) {
            var id = GetAttribute(item, "id");
            var href = GetAttribute(item, "href");
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(href)) continue;

            var model = new ManifestItem {
                Id = id,
                Href = href,
                FullPath = ResolveRelativePath(opfPath, href),
                MediaType = GetAttribute(item, "media-type"),
                Properties = GetAttribute(item, "properties")
            };
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

                package.Spine.Add(new SpineItem {
                    IdRef = idRef,
                    SpineIndex = index,
                    IsLinear = isLinear
                });
            }
        }

        return package;
    }

    private static Dictionary<string, string> BuildNavigationTitleMap(Dictionary<string, ZipArchiveEntry> entryIndex, EpubPackage? package, List<string> warnings) {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (package == null) return map;

        if (!string.IsNullOrWhiteSpace(package.NavDocumentPath)) {
            TryReadNavigationDocument(entryIndex, package.NavDocumentPath!, map, warnings);
        }
        if (!string.IsNullOrWhiteSpace(package.NcxPath)) {
            TryReadNcxDocument(entryIndex, package.NcxPath!, map, warnings);
        }

        return map;
    }

    private static void TryReadNavigationDocument(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        string navPath,
        Dictionary<string, string> map,
        List<string> warnings) {
        if (!entryIndex.TryGetValue(navPath, out var navEntry)) return;

        var navContent = ReadEntryText(navEntry);
        if (!TryParseXml(navContent, out var navDocument) || navDocument == null) {
            warnings.Add($"EPUB navigation document '{navPath}' could not be parsed.");
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
        List<string> warnings) {
        if (!entryIndex.TryGetValue(ncxPath, out var ncxEntry)) return;

        var ncxContent = ReadEntryText(ncxEntry);
        if (!TryParseXml(ncxContent, out var ncxDocument) || ncxDocument == null) {
            warnings.Add($"EPUB NCX document '{ncxPath}' could not be parsed.");
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

    private static List<ChapterCandidate> BuildChapterCandidates(
        ZipArchive archive,
        Dictionary<string, ZipArchiveEntry> entryIndex,
        EpubPackage? package,
        EpubReadOptions options,
        List<string> warnings) {
        var candidates = new List<ChapterCandidate>();
        var seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (package != null && options.PreferSpineOrder && package.Spine.Count > 0) {
            foreach (var spineItem in package.Spine.OrderBy(s => s.SpineIndex)) {
                if (!options.IncludeNonLinearSpineItems && !spineItem.IsLinear) {
                    continue;
                }

                if (!package.Manifest.TryGetValue(spineItem.IdRef, out var manifestItem)) {
                    warnings.Add($"EPUB spine idref '{spineItem.IdRef}' does not exist in manifest.");
                    continue;
                }

                if (!IsChapterManifestItem(manifestItem)) {
                    continue;
                }

                if (!entryIndex.TryGetValue(manifestItem.FullPath, out var chapterEntry)) {
                    warnings.Add($"EPUB manifest item '{manifestItem.FullPath}' referenced by spine was not found in archive.");
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
                    IsLinear = spineItem.IsLinear
                });
            }
        }

        var shouldFallbackScan = candidates.Count == 0 && options.FallbackToHtmlScan;
        if (!options.PreferSpineOrder) {
            shouldFallbackScan = true;
        }

        if (shouldFallbackScan) {
            IEnumerable<ZipArchiveEntry> scanEntries = archive.Entries.Where(e => IsChapterEntry(e.FullName));
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
                    IsLinear = null
                });
            }
        }

        return candidates;
    }

    private static Dictionary<string, ManifestItem> BuildManifestByPath(EpubPackage? package) {
        var map = new Dictionary<string, ManifestItem>(StringComparer.OrdinalIgnoreCase);
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

    private static string ReadEntryText(ZipArchiveEntry entry) {
        using var entryStream = entry.Open();
        using var reader = new StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 16 * 1024, leaveOpen: false);
        return reader.ReadToEnd();
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

    private static bool IsName(XElement element, string expectedLocalName) {
        return string.Equals(element.Name.LocalName, expectedLocalName, StringComparison.OrdinalIgnoreCase);
    }

    private static string GetAttribute(XElement element, string attributeName) {
        var attr = element.Attributes().FirstOrDefault(a => string.Equals(a.Name.LocalName, attributeName, StringComparison.OrdinalIgnoreCase));
        return attr?.Value ?? string.Empty;
    }

    private static string? TryGetFirstElementValue(XElement container, string localName) {
        foreach (var element in container.Elements()) {
            if (!string.Equals(element.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase)) continue;

            var normalized = NormalizeWhitespace(element.Value);
            return normalized.Length == 0 ? null : normalized;
        }

        return null;
    }

    private static bool ContainsSpaceSeparatedToken(string? value, string token) {
        var source = value ?? string.Empty;
        if (string.IsNullOrWhiteSpace(source)) return false;
        var parts = source.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var part in parts) {
            if (string.Equals(part, token, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static string ResolveRelativePath(string basePath, string relativeOrAbsolute) {
        var normalizedBase = NormalizePath(basePath);
        var normalizedRelative = NormalizePath(RemoveFragmentAndQuery(relativeOrAbsolute));
        if (normalizedRelative.Length == 0) return string.Empty;

        if (normalizedRelative.Length >= 2 && normalizedRelative[1] == ':') {
            return normalizedRelative;
        }

        if (normalizedRelative.StartsWith("/", StringComparison.Ordinal)) {
            return CollapsePathSegments(normalizedRelative.TrimStart('/'));
        }

        var baseDirectory = string.Empty;
        var lastSlash = normalizedBase.LastIndexOf('/');
        if (lastSlash >= 0) {
            baseDirectory = normalizedBase.Substring(0, lastSlash);
        }

        var combined = baseDirectory.Length == 0
            ? normalizedRelative
            : baseDirectory + "/" + normalizedRelative;
        return CollapsePathSegments(combined);
    }

    private static string RemoveFragmentAndQuery(string value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        var trimmed = value.Trim();
        var hash = trimmed.IndexOf('#');
        if (hash >= 0) {
            trimmed = trimmed.Substring(0, hash);
        }

        var question = trimmed.IndexOf('?');
        if (question >= 0) {
            trimmed = trimmed.Substring(0, question);
        }

        if (trimmed.Length == 0) return string.Empty;

        try {
            return Uri.UnescapeDataString(trimmed);
        } catch {
            return trimmed;
        }
    }

    private static string CollapsePathSegments(string path) {
        var segments = path.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        var stack = new List<string>(segments.Length);

        foreach (var segment in segments) {
            if (segment == ".") continue;
            if (segment == "..") {
                if (stack.Count > 0) {
                    stack.RemoveAt(stack.Count - 1);
                }
                continue;
            }

            stack.Add(segment);
        }

        return string.Join("/", stack);
    }

    private static string NormalizePath(string path) {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;
        return path.Replace('\\', '/').Trim();
    }

    private static string NormalizeWhitespace(string value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        var sb = new StringBuilder(value.Length);
        bool previousWasWhitespace = false;
        foreach (var ch in value) {
            if (char.IsWhiteSpace(ch)) {
                if (!previousWasWhitespace) {
                    sb.Append(' ');
                    previousWasWhitespace = true;
                }
            } else {
                sb.Append(ch);
                previousWasWhitespace = false;
            }
        }

        return sb.ToString().Trim();
    }

    private static EpubReadOptions Normalize(EpubReadOptions? options) {
        var source = options ?? new EpubReadOptions();
        var normalized = new EpubReadOptions {
            MaxChapters = source.MaxChapters,
            MaxChapterBytes = source.MaxChapterBytes,
            IncludeRawHtml = source.IncludeRawHtml,
            DeterministicOrder = source.DeterministicOrder,
            PreferSpineOrder = source.PreferSpineOrder,
            IncludeNonLinearSpineItems = source.IncludeNonLinearSpineItems,
            FallbackToHtmlScan = source.FallbackToHtmlScan
        };

        if (normalized.MaxChapters < 1) normalized.MaxChapters = 1;
        if (normalized.MaxChapterBytes.HasValue && normalized.MaxChapterBytes.Value < 1) {
            normalized.MaxChapterBytes = 1;
        }

        return normalized;
    }

    private sealed class EpubPackage {
        public string OpfPath { get; set; } = string.Empty;
        public string? Title { get; set; }
        public string? Identifier { get; set; }
        public string? Language { get; set; }
        public string? Creator { get; set; }
        public Dictionary<string, ManifestItem> Manifest { get; } = new Dictionary<string, ManifestItem>(StringComparer.OrdinalIgnoreCase);
        public List<SpineItem> Spine { get; } = new List<SpineItem>();
        public string? NavDocumentPath { get; set; }
        public string? NcxPath { get; set; }
    }

    private sealed class ManifestItem {
        public string Id { get; set; } = string.Empty;
        public string Href { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public string MediaType { get; set; } = string.Empty;
        public string Properties { get; set; } = string.Empty;
    }

    private sealed class SpineItem {
        public string IdRef { get; set; } = string.Empty;
        public int SpineIndex { get; set; }
        public bool IsLinear { get; set; }
    }

    private sealed class ChapterCandidate {
        public ZipArchiveEntry Entry { get; set; } = null!;
        public string? ManifestId { get; set; }
        public string? MediaType { get; set; }
        public int? SpineIndex { get; set; }
        public bool? IsLinear { get; set; }
    }
}
