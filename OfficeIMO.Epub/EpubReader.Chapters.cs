namespace OfficeIMO.Epub;

#if NET8_0_OR_GREATER
using System.Buffers;
#endif

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
                if (seenPaths.Contains(chapterPath)) {
                    continue;
                }

                seenPaths.Add(chapterPath);
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

        var shouldFallbackScan = candidates.Count == 0 && options.FallbackToHtmlScan;
        if (!options.PreferSpineOrder) {
            shouldFallbackScan = true;
        }

        if (shouldFallbackScan) {
            IEnumerable<KeyValuePair<string, ZipArchiveEntry>> scanEntries = entryIndex
                .Where(entry => IsChapterEntry(entry.Key));
            if (options.DeterministicOrder) {
                scanEntries = scanEntries.OrderBy(entry => entry.Key, StringComparer.Ordinal);
            }

            var manifestByPath = BuildManifestByPath(package);
            foreach (KeyValuePair<string, ZipArchiveEntry> indexedEntry in scanEntries) {
                string chapterPath = indexedEntry.Key;
                if (seenPaths.Contains(chapterPath)) continue;

                manifestByPath.TryGetValue(chapterPath, out var manifestItem);
                seenPaths.Add(chapterPath);
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
        byte[] data = ReadEntryBytesExact(entry, maxBytes);
        if (data.Length >= 4) {
            if (data[0] == 0x00 && data[1] == 0x00 && data[2] == 0xFE && data[3] == 0xFF) {
                return BigEndianUtf32.GetString(data, 4, data.Length - 4);
            }
            if (data[0] == 0xFF && data[1] == 0xFE && data[2] == 0x00 && data[3] == 0x00) {
                return Encoding.UTF32.GetString(data, 4, data.Length - 4);
            }
        }
        if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) {
            return Encoding.UTF8.GetString(data, 3, data.Length - 3);
        }
        if (data.Length >= 2) {
            if (data[0] == 0xFE && data[1] == 0xFF) {
                return Encoding.BigEndianUnicode.GetString(data, 2, data.Length - 2);
            }
            if (data[0] == 0xFF && data[1] == 0xFE) {
                return Encoding.Unicode.GetString(data, 2, data.Length - 2);
            }
        }
        return Encoding.UTF8.GetString(data);
    }

    private static byte[] ReadEntryBytes(ZipArchiveEntry entry, long maxBytes) {
        return ReadEntryBytesExact(entry, maxBytes);
    }

    private static byte[] ReadEntryBytesExact(ZipArchiveEntry entry, long? maxBytes) {
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

    private static bool TryReadChapterMarkup(string content, out ChapterMarkupInfo chapter) {
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
            bool sawBody = false;
            bool hasVisibleText = false;
            bool pendingVisibleSpace = false;
            int visibleTextLength = 0;
            bool hasStructuredContent = false;
            string? baseHref = null;

            while (reader.Read()) {
                switch (reader.NodeType) {
                    case XmlNodeType.Element:
                        string localName = reader.LocalName;
                        if (!sawBody && localName.Equals("body", StringComparison.OrdinalIgnoreCase)) {
                            sawBody = true;
                            bodyDepth = reader.Depth;
                            visibleTextLength = 0;
                            hasVisibleText = false;
                            pendingVisibleSpace = false;
                        }
                        if (excludedTextDepth < 0 &&
                            (localName.Equals("script", StringComparison.OrdinalIgnoreCase) ||
                             localName.Equals("style", StringComparison.OrdinalIgnoreCase))) {
                            excludedTextDepth = reader.Depth;
                        }
                        if (reader.Depth > 0 && title == null && localName.Equals("title", StringComparison.OrdinalIgnoreCase)) {
                            title = new StringBuilder();
                            titleDepth = reader.Depth;
                        }
                        if (reader.Depth > 0 && heading == null &&
                            (localName.Equals("h1", StringComparison.OrdinalIgnoreCase) ||
                             localName.Equals("h2", StringComparison.OrdinalIgnoreCase))) {
                            heading = new StringBuilder();
                            headingDepth = reader.Depth;
                        }
                        if (reader.Depth > 0 && baseHref == null && localName.Equals("base", StringComparison.OrdinalIgnoreCase)) {
                            baseHref = NullIfWhiteSpace(GetAttribute(reader, "href"));
                        }
                        if (reader.Depth > 0 && !hasStructuredContent && IsStructuredChapterElement(localName)) {
                            hasStructuredContent = true;
                        }
                        if (reader.IsEmptyElement) {
                            if (reader.Depth == excludedTextDepth) excludedTextDepth = -1;
                            if (reader.Depth == titleDepth) titleDepth = -1;
                            if (reader.Depth == headingDepth) headingDepth = -1;
                            if (reader.Depth == bodyDepth) bodyDepth = -1;
                        }
                        break;

                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.SignificantWhitespace:
                    case XmlNodeType.Whitespace:
                        if (titleDepth >= 0 && reader.Depth > titleDepth) title!.Append(reader.Value);
                        if (headingDepth >= 0 && reader.Depth > headingDepth) heading!.Append(reader.Value);
                        bool withinSelectedScope = sawBody
                            ? bodyDepth >= 0 && reader.Depth > bodyDepth
                            : reader.Depth > 0;
                        bool hasExcludedDirectParent = excludedTextDepth >= 0 && reader.Depth == excludedTextDepth + 1;
                        if (withinSelectedScope && !hasExcludedDirectParent && !string.IsNullOrWhiteSpace(reader.Value)) {
                            AppendNormalizedVisibleText(
                                visibleText,
                                ref visibleTextLength,
                                reader.Value,
                                ref hasVisibleText,
                                ref pendingVisibleSpace);
                            pendingVisibleSpace = hasVisibleText;
                        }
                        break;

                    case XmlNodeType.EndElement:
                        if (reader.Depth == excludedTextDepth) excludedTextDepth = -1;
                        if (reader.Depth == titleDepth) titleDepth = -1;
                        if (reader.Depth == headingDepth) headingDepth = -1;
                        if (reader.Depth == bodyDepth) bodyDepth = -1;
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
        } catch {
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
            if (reader.LocalName.Equals(attributeName, StringComparison.OrdinalIgnoreCase)) {
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

    private static readonly Encoding BigEndianUtf32 = new UTF32Encoding(bigEndian: true, byteOrderMark: true);

}
