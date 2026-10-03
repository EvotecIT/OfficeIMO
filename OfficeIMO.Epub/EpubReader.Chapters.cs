namespace OfficeIMO.Epub;

#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.Threading;

internal static partial class EpubReader {
    private static List<ChapterCandidate> BuildChapterCandidates(
        Dictionary<string, ZipArchiveEntry> entryIndex,
        EpubPackage? package,
        EpubReadOptions options,
        EpubDiagnosticCollector diagnostics,
        CancellationToken cancellationToken,
        out EpubReadSummary readSummary) {
        var candidates = new List<ChapterCandidate>();
        var fallbackSelections = new Dictionary<string, ManifestItem?>(StringComparer.Ordinal);
        readSummary = new EpubReadSummary { IsSpineBased = package != null && package.Spine.Count > 0 };

        if (package != null && readSummary.IsSpineBased) {
            foreach (var spineItem in package.Spine.OrderBy(s => s.SpineIndex)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!options.IncludeNonLinearSpineItems && !spineItem.IsLinear) {
                    continue;
                }
                readSummary.RequestedChapterCount++;

                if (!package.Manifest.TryGetValue(spineItem.IdRef, out var manifestItem)) {
                    diagnostics.Warning(
                        "epub.spine.manifest-id-missing",
                        $"EPUB spine idref '{spineItem.IdRef}' does not exist in manifest.",
                        package.OpfPath);
                    continue;
                }

                ManifestItem? selected = ResolveChapterResource(package, manifestItem, fallbackSelections, diagnostics, cancellationToken);
                if (selected == null) continue;
                manifestItem = selected;

                if (manifestItem.IsRemote) {
                    diagnostics.Warning(
                        "epub.spine.remote-resource",
                        $"Skipped remote spine resource '{manifestItem.RemoteUri}' because remote content is not fetched.",
                        manifestItem.RemoteUri);
                    continue;
                }

                if (!entryIndex.TryGetValue(manifestItem.FullPath, out var chapterEntry)) {
                    diagnostics.Warning(
                        "epub.spine.resource-missing",
                        $"EPUB manifest item '{manifestItem.FullPath}' referenced by spine was not found in archive.",
                        manifestItem.FullPath);
                    continue;
                }

                string chapterPath = manifestItem.FullPath;
                candidates.Add(new ChapterCandidate {
                    Entry = chapterEntry,
                    Path = chapterPath,
                    ManifestId = manifestItem.Id,
                    MediaType = manifestItem.MediaType,
                    SpineIndex = spineItem.SpineIndex,
                    IsLinear = spineItem.IsLinear,
                    RenditionLayout = spineItem.RenditionLayout
                });
            }
        }

        bool shouldFallbackScan = !readSummary.IsSpineBased && options.FallbackToHtmlScan;

        if (shouldFallbackScan) {
            readSummary.UsedFallbackScan = true;
            diagnostics.Warning("epub.chapter.fallback-scan",
                "Chapters were recovered by scanning the archive because no usable spine was declared; publication completeness cannot be established.",
                package?.OpfPath);
            IEnumerable<KeyValuePair<string, ZipArchiveEntry>> scanEntries = entryIndex
                .Where(entry => IsChapterEntry(entry.Key));
            if (options.DeterministicOrder) {
                scanEntries = scanEntries.OrderBy(entry => entry.Key, StringComparer.Ordinal);
            }

            var manifestByPath = BuildManifestByPath(package);
            foreach (KeyValuePair<string, ZipArchiveEntry> indexedEntry in scanEntries) {
                cancellationToken.ThrowIfCancellationRequested();
                string chapterPath = indexedEntry.Key;
                manifestByPath.TryGetValue(chapterPath, out var manifestItem);
                if (manifestItem != null && ContainsSpaceSeparatedToken(manifestItem.Properties, "nav")) continue;
                readSummary.RequestedChapterCount++;
                candidates.Add(new ChapterCandidate {
                    Entry = indexedEntry.Value,
                    Path = chapterPath,
                    ManifestId = manifestItem?.Id,
                    MediaType = manifestItem?.MediaType,
                    SpineIndex = null,
                    IsLinear = null,
                    RenditionLayout = package?.RenditionLayout
                });
            }
        }

        if (options.PreferSpineOrder || candidates.Count == 0) return candidates;
        if (options.DeterministicOrder)
            return candidates.OrderBy(candidate => candidate.Path, StringComparer.Ordinal).ToList();
        var archiveOrder = new Dictionary<ZipArchiveEntry, int>();
        foreach (ZipArchiveEntry entry in candidates[0].Entry.Archive.Entries)
            archiveOrder.Add(entry, archiveOrder.Count);
        return candidates.OrderBy(candidate => archiveOrder[candidate.Entry]).ToList();
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
        if (string.Equals(item.MediaType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(item.MediaType, "application/xhtml+xml", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(item.MediaType, "text/html", StringComparison.OrdinalIgnoreCase)) return true;
        return string.IsNullOrWhiteSpace(item.MediaType) && IsChapterEntry(item.FullPath);
    }

    private static bool IsChapterEntry(string? fullName) {
        if (string.IsNullOrWhiteSpace(fullName)) return false;
        var normalized = NormalizePath(fullName!);
        if (normalized.EndsWith("/", StringComparison.Ordinal)) return false;

        var ext = Path.GetExtension(normalized).ToLowerInvariant();
        return ext == ".xhtml" || ext == ".html" || ext == ".htm";
    }

    private static string ReadEntryText(
        ZipArchiveEntry entry,
        long? maxBytes,
        CancellationToken cancellationToken = default) {
        byte[] data = ReadEntryBytesExact(entry, maxBytes, cancellationToken);
        if (data.Length >= 4) {
            if (data[0] == 0x00 && data[1] == 0x00 && data[2] == 0xFE && data[3] == 0xFF) {
                return StrictBigEndianUtf32.GetString(data, 4, data.Length - 4);
            }
            if (data[0] == 0xFF && data[1] == 0xFE && data[2] == 0x00 && data[3] == 0x00) {
                return StrictUtf32.GetString(data, 4, data.Length - 4);
            }
        }
        if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) {
            return StrictUtf8.GetString(data, 3, data.Length - 3);
        }
        if (data.Length >= 2) {
            if (data[0] == 0xFE && data[1] == 0xFF) {
                return StrictBigEndianUtf16.GetString(data, 2, data.Length - 2);
            }
            if (data[0] == 0xFF && data[1] == 0xFE) {
                return StrictUtf16.GetString(data, 2, data.Length - 2);
            }
        }
        if (data.Length >= 4 && data[0] == 0 && data[1] == '<' && data[2] == 0) return StrictBigEndianUtf16.GetString(data);
        if (data.Length >= 4 && data[0] == '<' && data[1] == 0 && data[3] == 0) return StrictUtf16.GetString(data);
        return StrictUtf8.GetString(data);
    }

    private static byte[] ReadEntryBytes(
        ZipArchiveEntry entry,
        long maxBytes,
        CancellationToken cancellationToken = default) {
        return ReadEntryBytesExact(entry, maxBytes, cancellationToken);
    }

    private static byte[] ReadEntryBytesExact(
        ZipArchiveEntry entry,
        long? maxBytes,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (maxBytes.HasValue && entry.Length > maxBytes.Value) {
            throw new InvalidDataException($"EPUB entry '{entry.FullName}' exceeds the configured maximum size ({maxBytes.Value} bytes).");
        }
        if (entry.Length > int.MaxValue) {
            throw new InvalidDataException($"EPUB entry '{entry.FullName}' exceeds the supported in-memory size.");
        }

        using Stream entryStream = entry.Open();
        if (entry.Length == 0) {
            if (entryStream.ReadByte() >= 0) {
                throw new InvalidDataException(
                    $"EPUB entry '{entry.FullName}' expanded beyond its declared uncompressed size.");
            }
            return Array.Empty<byte>();
        }

        const int bufferSize = 81920;
        int initialCapacity = checked((int)Math.Min(entry.Length, bufferSize));
        using var output = new MemoryStream(initialCapacity);
#if NET8_0_OR_GREATER
        byte[] buffer = ArrayPool<byte>.Shared.Rent(initialCapacity);
#else
        byte[] buffer = new byte[initialCapacity];
#endif
        try {
            long total = 0;
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = entryStream.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                if (read > entry.Length - total) {
                    throw new InvalidDataException(
                        $"EPUB entry '{entry.FullName}' expanded beyond its declared uncompressed size.");
                }
                if (maxBytes.HasValue && read > maxBytes.Value - total) {
                    throw new InvalidDataException(
                        $"EPUB entry '{entry.FullName}' exceeds the configured maximum size ({maxBytes.Value} bytes).");
                }
                output.Write(buffer, 0, read);
                total += read;
            }
            return output.ToArray();
        } finally {
#if NET8_0_OR_GREATER
            ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
#endif
        }
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

    private static bool TryParseEntryXml(
        ZipArchiveEntry entry,
        long maxBytes,
        CancellationToken cancellationToken,
        out XDocument? document) =>
        TryParseXml(ReadEntryBytesExact(entry, maxBytes, cancellationToken), out document);

    private static bool TryParseXml(byte[] content, out XDocument? document) {
        document = null;
        if (content.Length == 0) return false;

        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Ignore,
                XmlResolver = null
            };

            using var input = new MemoryStream(content, writable: false);
            using var xmlReader = XmlReader.Create(input, settings);
            document = XDocument.Load(xmlReader, LoadOptions.PreserveWhitespace);
            return true;
        } catch {
            return false;
        }
    }

    private static bool TryReadChapterMarkup(string content, out ChapterMarkupInfo chapter, CancellationToken cancellationToken) {
        chapter = ChapterMarkupInfo.Empty;
        if (string.IsNullOrWhiteSpace(content)) return false;

#if NET8_0_OR_GREATER
        char[] visibleText = ArrayPool<char>.Shared.Rent(content.Length);
#else
        char[] visibleText = new char[content.Length];
#endif
        try {
            var settings = new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Ignore,
                XmlResolver = null
            };
            using var stringReader = new StringReader(content);
            using XmlReader reader = XmlReader.Create(stringReader, settings);
            StringBuilder? title = null;
            StringBuilder? heading = null;
            int bodyDepth = -1;
            int excludedTextDepth = -1;
            int titleDepth = -1;
            int headingDepth = -1;
            int svgTextDepth = -1;
            bool isSvg = false;
            bool sawBody = false;
            bool hasVisibleText = false;
            bool pendingVisibleSpace = false;
            int visibleTextLength = 0;
            bool hasStructuredContent = false;
            string? baseHref = null;

            while (reader.Read()) {
                cancellationToken.ThrowIfCancellationRequested();
                switch (reader.NodeType) {
                    case XmlNodeType.Element:
                        string localName = reader.LocalName;
                        bool isHtmlElement = IsXhtmlNamespace(reader.NamespaceURI);
                        bool isSvgElement = reader.NamespaceURI == SvgNamespaceUri;
                        if (reader.Depth == 0) {
                            isSvg = isSvgElement && localName == "svg";
                            if (!isSvg && !(isHtmlElement && localName.Equals("html", StringComparison.OrdinalIgnoreCase))) return false;
                            hasStructuredContent = isSvg;
                        }
                        if (!isSvg && !sawBody && reader.Depth == 1 && isHtmlElement &&
                            localName.Equals("body", StringComparison.OrdinalIgnoreCase)) {
                            sawBody = true;
                            bodyDepth = reader.Depth;
                            visibleTextLength = 0;
                            hasVisibleText = false;
                            pendingVisibleSpace = false;
                        }
                        if (excludedTextDepth < 0 && (isHtmlElement || isSvgElement) &&
                            (localName.Equals("script", StringComparison.OrdinalIgnoreCase) ||
                             localName.Equals("style", StringComparison.OrdinalIgnoreCase))) {
                            excludedTextDepth = reader.Depth;
                        }
                        if (reader.Depth > 0 && title == null && (isHtmlElement || isSvgElement) && localName.Equals("title", StringComparison.OrdinalIgnoreCase)) {
                            title = new StringBuilder();
                            titleDepth = reader.Depth;
                        }
                        if (reader.Depth > 0 && heading == null && isHtmlElement &&
                            (localName.Equals("h1", StringComparison.OrdinalIgnoreCase) ||
                             localName.Equals("h2", StringComparison.OrdinalIgnoreCase))) {
                            heading = new StringBuilder();
                            headingDepth = reader.Depth;
                        }
                        if (!isSvg && reader.Depth > 0 && baseHref == null && isHtmlElement && localName.Equals("base", StringComparison.OrdinalIgnoreCase)) {
                            baseHref = NullIfWhiteSpace(GetAttribute(reader, "href"));
                        }
                        if (reader.Depth > 0 && !hasStructuredContent && (isHtmlElement || isSvgElement) && IsStructuredChapterElement(localName)) {
                            hasStructuredContent = true;
                        }
                        if (isSvgElement && svgTextDepth < 0 && localName == "text") svgTextDepth = reader.Depth;
                        if (isHtmlElement && IsTextBoundaryElement(localName)) pendingVisibleSpace = hasVisibleText;
                        if (reader.IsEmptyElement) {
                            if (reader.Depth == excludedTextDepth) excludedTextDepth = -1;
                            if (reader.Depth == titleDepth) titleDepth = -1;
                            if (reader.Depth == headingDepth) headingDepth = -1;
                            if (reader.Depth == bodyDepth) bodyDepth = -1;
                            if (reader.Depth == svgTextDepth) svgTextDepth = -1;
                        }
                        break;

                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.SignificantWhitespace:
                    case XmlNodeType.Whitespace:
                        if (titleDepth >= 0 && reader.Depth > titleDepth) title!.Append(reader.Value);
                        if (headingDepth >= 0 && reader.Depth > headingDepth) heading!.Append(reader.Value);
                        bool withinSelectedScope = isSvg ? svgTextDepth >= 0 && reader.Depth > svgTextDepth : sawBody
                            ? bodyDepth >= 0 && reader.Depth > bodyDepth
                            : reader.Depth > 0;
                        bool isExcluded = excludedTextDepth >= 0 && reader.Depth > excludedTextDepth;
                        if (withinSelectedScope && !isExcluded) {
                            AppendNormalizedVisibleText(
                                visibleText,
                                ref visibleTextLength,
                                reader.Value,
                                ref hasVisibleText,
                                ref pendingVisibleSpace);
                        }
                        break;

                    case XmlNodeType.EndElement:
                        if (IsXhtmlNamespace(reader.NamespaceURI) && IsTextBoundaryElement(reader.LocalName)) pendingVisibleSpace = hasVisibleText;
                        if (reader.Depth == excludedTextDepth) excludedTextDepth = -1;
                        if (reader.Depth == titleDepth) titleDepth = -1;
                        if (reader.Depth == headingDepth) headingDepth = -1;
                        if (reader.Depth == bodyDepth) bodyDepth = -1;
                        if (reader.Depth == svgTextDepth) {
                            svgTextDepth = -1;
                            pendingVisibleSpace = hasVisibleText;
                        }
                        break;
                }
            }

            chapter = new ChapterMarkupInfo(
                visibleTextLength == 0 ? string.Empty : new string(visibleText, 0, visibleTextLength),
                NormalizeOptional(title),
                NormalizeOptional(heading),
                baseHref,
                hasStructuredContent);
            return true;
        } catch (XmlException) {
            chapter = ChapterMarkupInfo.Empty;
            return false;
        } finally {
#if NET8_0_OR_GREATER
            ArrayPool<char>.Shared.Return(visibleText);
#endif
        }
    }

    private static void AppendNormalizedVisibleText(
        char[] destination,
        ref int length,
        string value,
        ref bool hasText,
        ref bool pendingSpace) {
        foreach (char character in value) {
            if (char.IsWhiteSpace(character)) {
                pendingSpace = hasText;
                continue;
            }
            if (pendingSpace && hasText) destination[length++] = ' ';
            destination[length++] = character;
            hasText = true;
            pendingSpace = false;
        }
    }

    private static string? NormalizeOptional(StringBuilder? value) {
        if (value == null || value.Length == 0) return null;
        string normalized = NormalizeWhitespace(value.ToString());
        return normalized.Length == 0 ? null : normalized;
    }

    private static string GetAttribute(XmlReader reader, string attributeName) {
        if (!reader.HasAttributes) return string.Empty;
        while (reader.MoveToNextAttribute()) {
            if (reader.NamespaceURI.Length == 0 && reader.LocalName.Equals(attributeName, StringComparison.OrdinalIgnoreCase)) {
                string value = reader.Value;
                reader.MoveToElement();
                return value;
            }
        }
        reader.MoveToElement();
        return string.Empty;
    }

    private static bool IsStructuredChapterElement(string localName) =>
        localName.Equals("img", StringComparison.OrdinalIgnoreCase) ||
        localName.Equals("picture", StringComparison.OrdinalIgnoreCase) ||
        localName.Equals("svg", StringComparison.OrdinalIgnoreCase) ||
        localName.Equals("table", StringComparison.OrdinalIgnoreCase) ||
        localName.Equals("form", StringComparison.OrdinalIgnoreCase) ||
        localName.Equals("input", StringComparison.OrdinalIgnoreCase) ||
        localName.Equals("select", StringComparison.OrdinalIgnoreCase) ||
        localName.Equals("textarea", StringComparison.OrdinalIgnoreCase) ||
        localName.Equals("audio", StringComparison.OrdinalIgnoreCase) ||
        localName.Equals("video", StringComparison.OrdinalIgnoreCase) ||
        localName.Equals("object", StringComparison.OrdinalIgnoreCase) ||
        localName.Equals("canvas", StringComparison.OrdinalIgnoreCase);

    private static string? ResolveChapterTitle(ChapterMarkupInfo chapter, Dictionary<string, string> navTitleMap, string chapterPath) {
        if (navTitleMap.TryGetValue(chapterPath, out var navTitle) && !string.IsNullOrWhiteSpace(navTitle)) {
            return navTitle;
        }
        return chapter.Title ?? chapter.Heading;
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

    private static bool IsTextBoundaryElement(string name) => TextBoundaryElements.Contains(name);

    private static readonly HashSet<string> TextBoundaryElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "address", "article", "aside", "blockquote", "br", "caption", "dd", "div", "dl", "dt", "fieldset",
        "figcaption", "figure", "footer", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr", "li", "main",
        "nav", "ol", "p", "pre", "section", "table", "td", "th", "tr", "ul"
    };

    private static readonly Encoding StrictUtf8 = new UTF8Encoding(false, true);
    private static readonly Encoding StrictUtf16 = new UnicodeEncoding(false, true, true);
    private static readonly Encoding StrictBigEndianUtf16 = new UnicodeEncoding(true, true, true);
    private static readonly Encoding StrictUtf32 = new UTF32Encoding(false, true, true);
    private static readonly Encoding StrictBigEndianUtf32 = new UTF32Encoding(true, true, true);

}
