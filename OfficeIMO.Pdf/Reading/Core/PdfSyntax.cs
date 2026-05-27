using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal static class PdfSyntax {
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);
#if NET8_0_OR_GREATER
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+(\d+)\s+obj", RegexOptions.Compiled | RegexOptions.NonBacktracking, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"<<(.*?)>>\s*stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.NonBacktracking, RegexTimeout);
#else
    private static readonly Regex ObjRegex = new Regex(@"(\d+)\s+(\d+)\s+obj", RegexOptions.Compiled, RegexTimeout);
    private static readonly Regex StreamRegex = new Regex(@"<<(.*?)>>\s*stream\r?\n([\s\S]*?)\r?\nendstream", RegexOptions.Compiled | RegexOptions.Singleline, RegexTimeout);
#endif
    private static readonly Regex TrailerRootRegex = new Regex(@"/Root\s+(\d+)\s+(\d+)\s+R", RegexOptions.Compiled, RegexTimeout);

    internal static (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw) ParseObjects(byte[] pdf) {
        string text = PdfEncoding.Latin1GetString(pdf);
        var map = new Dictionary<int, PdfIndirectObject>();
        var parsedOffsets = new Dictionary<int, int>();
        var streamLocations = new List<(int Id, int Generation, int DataStart)>();
        var matches = ObjRegex.Matches(text);
        for (int i = 0; i < matches.Count; i++) {
            int id = int.Parse(matches[i].Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            int gen = int.Parse(matches[i].Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
            int start = matches[i].Index;
            int bodyStart = matches[i].Index + matches[i].Length;
            int end = FindObjectEnd(text, start);
            if (end < 0) end = (i + 1 < matches.Count) ? matches[i + 1].Index : text.Length;

            // Extract dictionary (balanced << >>) within object bounds
            int dictStart = text.IndexOf("<<", start, end - start, System.StringComparison.Ordinal);
            if (dictStart >= 0) {
                int dictEnd = FindDictEnd(text, dictStart, end);
                if (dictEnd > dictStart) {
                    string dictText = SafeSlice(text, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000); // cap to 1 MB
                    PdfDictionary? dict;
                    try { dict = ParseDictionary(dictText); }
                    catch (Exception ex) when (ex is not OutOfMemoryException) { dict = null; }
                    if (dict is null) {
                        continue;
                    }

                    // Check for stream section; prefer dictionary /Length when available
                    int streamKw = IndexOfKeyword(text, "stream", dictEnd, end);
                    if (streamKw >= 0) {
                        int dataStart = SkipEOL(text, streamKw + 6, end);
                        streamLocations.Add((id, gen, dataStart));
                        // Try /Length first (inline number only)
                        int byteStart = dataStart;
                        int byteLen = -1;
                        TryGetResolvedLength(dict, map, out byteLen);
                        if (byteLen < 0) {
                            int endStream = IndexOfKeyword(text, "endstream", dataStart, end);
                            if (endStream > dataStart) byteLen = endStream - dataStart;
                        }
                        if (byteLen >= 0) {
                            if (byteStart >= 0 && byteLen >= 0 && byteStart + byteLen <= pdf.Length) {
                                var data = new byte[byteLen];
                                Buffer.BlockCopy(pdf, byteStart, data, 0, byteLen);
                                map[id] = new PdfIndirectObject(id, gen, new PdfStream(dict, data));
                                parsedOffsets[id] = start;
                                continue;
                            }
                        }
                    }
                    // No stream; store dictionary-only object
                    map[id] = new PdfIndirectObject(id, gen, dict);
                    parsedOffsets[id] = start;
                }
            }

            if (!map.ContainsKey(id)) {
                int bodyEnd = end;
                if (bodyEnd - 6 >= bodyStart && string.Equals(text.Substring(bodyEnd - 6, 6), "endobj", StringComparison.Ordinal)) {
                    bodyEnd -= 6;
                }

                string body = SafeSlice(text, bodyStart, bodyEnd - bodyStart, 1_000_000).Trim();
                var parsed = ParseTopLevelObject(body);
                if (parsed is not null) {
                    map[id] = new PdfIndirectObject(id, gen, parsed);
                    parsedOffsets[id] = start;
                }
            }
        }
        string trailerRaw = GetActiveTrailerRaw(text, map, parsedOffsets);
        ThrowIfEncrypted(trailerRaw);

        ResolveIndirectStreamLengths(map, pdf, streamLocations);
        var activeClassicObjectNumbers = new HashSet<int>();
        bool appliedXrefStreamEntries = ApplyClassicXrefEntries(map, pdf, parsedOffsets, activeClassicObjectNumbers, out bool appliedClassicEntries);
        appliedXrefStreamEntries = ApplyXrefStreamEntries(map, pdf, parsedOffsets) || appliedXrefStreamEntries;
        if (!appliedXrefStreamEntries) {
            // Compatibility fallback for simple parser-supported files whose compressed objects are only discoverable by scanning.
            ExpandObjectStreams(map, pdf, parsedOffsets, appliedClassicEntries ? activeClassicObjectNumbers : null);
        }

        ThrowIfEncryptedXrefStream(map);
        return (map, trailerRaw);
    }

    private static string GetActiveTrailerRaw(string text, Dictionary<int, PdfIndirectObject> map, Dictionary<int, int> parsedOffsets) {
        if (TryGetLatestStartXrefOffset(text, out int activeXrefOffset)) {
            if (TryGetClassicTrailerChainRaw(text, map, parsedOffsets, activeXrefOffset, out string trailerRaw)) {
                return trailerRaw;
            }

            if (TryGetXrefStreamTrailerChainRaw(text, map, parsedOffsets, activeXrefOffset, out trailerRaw)) {
                return trailerRaw;
            }
        }

        int trailerIdx = text.LastIndexOf("trailer", StringComparison.OrdinalIgnoreCase);
        return trailerIdx >= 0 ? text.Substring(trailerIdx) : string.Empty;
    }

    private static bool TryGetXrefStreamTrailerChainRaw(
        string text,
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int activeXrefOffset,
        out string trailerRaw) {
        trailerRaw = string.Empty;
        var byOffset = new Dictionary<int, PdfDictionary>();
        foreach (var entry in map.Values) {
            if (!parsedOffsets.TryGetValue(entry.ObjectNumber, out int offset)) {
                continue;
            }

            PdfDictionary? dictionary = entry.Value is PdfStream stream ? stream.Dictionary : entry.Value as PdfDictionary;
            if (dictionary?.Get<PdfName>("Type")?.Name == "XRef") {
                byOffset[offset] = dictionary;
            }
        }

        var trailers = new List<string>();
        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (byOffset.TryGetValue(currentOffset, out PdfDictionary? dictionary) &&
            visited.Add(currentOffset) &&
            trailers.Count < 64) {
            trailers.Add(BuildXrefStreamTrailerRaw(dictionary));
            if (dictionary.Get<PdfNumber>("Prev") is not PdfNumber previous ||
                previous.Value < 0 ||
                previous.Value > int.MaxValue) {
                break;
            }

            currentOffset = (int)Math.Floor(previous.Value);
        }

        if (trailers.Count > 0 &&
            TryGetClassicTrailerChainRaw(text, map, parsedOffsets, currentOffset, out string classicTrailerRaw)) {
            trailers.Add(classicTrailerRaw);
        }

        if (trailers.Count == 0) {
            return false;
        }

        trailerRaw = string.Join("\n", trailers);
        return true;
    }

    private static string BuildXrefStreamTrailerRaw(PdfDictionary dictionary) {
        var parts = new List<string>();
        AppendTrailerEntry(parts, dictionary, "Size");
        AppendTrailerEntry(parts, dictionary, "Root");
        AppendTrailerEntry(parts, dictionary, "Info");
        AppendTrailerEntry(parts, dictionary, "ID");
        AppendTrailerEntry(parts, dictionary, "Encrypt");
        AppendTrailerEntry(parts, dictionary, "Prev");
        return "trailer\n<< " + string.Join(" ", parts) + " >>";
    }

    private static void AppendTrailerEntry(List<string> parts, PdfDictionary dictionary, string key) {
        if (dictionary.Items.TryGetValue(key, out PdfObject? value) &&
            TryFormatTrailerValue(value, out string? formatted)) {
            parts.Add("/" + key + " " + formatted);
        }
    }

    private static bool TryFormatTrailerValue(PdfObject value, out string? formatted) {
        switch (value) {
            case PdfReference reference:
                formatted = reference.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " +
                    reference.Generation.ToString(System.Globalization.CultureInfo.InvariantCulture) + " R";
                return true;
            case PdfNumber number:
                formatted = number.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                return true;
            case PdfName name:
                formatted = "/" + name.Name;
                return true;
            case PdfStringObj text:
                formatted = "(" + text.Value.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)") + ")";
                return true;
            case PdfArray array:
                var items = new List<string>();
                foreach (PdfObject item in array.Items) {
                    if (!TryFormatTrailerValue(item, out string? itemText)) {
                        formatted = null;
                        return false;
                    }

                    if (itemText is null) {
                        formatted = null;
                        return false;
                    }

                    items.Add(itemText);
                }

                formatted = "[" + string.Join(" ", items) + "]";
                return true;
            case PdfNull:
                formatted = "null";
                return true;
            default:
                formatted = null;
                return false;
        }
    }

    private static bool TryGetClassicTrailerChainRaw(
        string text,
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int activeXrefOffset,
        out string trailerRaw) {
        trailerRaw = string.Empty;
        var trailers = new List<string>();
        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (visited.Add(currentOffset) &&
            trailers.Count < 64 &&
            TryParseClassicXrefTable(text, currentOffset, out _, out int? previousOffset, out string currentTrailerRaw, out int? xrefStreamOffset)) {
            if (!string.IsNullOrWhiteSpace(currentTrailerRaw)) {
                trailers.Add(currentTrailerRaw);
            }

            if (xrefStreamOffset.HasValue &&
                TryGetXrefStreamTrailerRawAtOffset(map, parsedOffsets, xrefStreamOffset.Value, out string xrefStreamTrailerRaw)) {
                trailers.Add(xrefStreamTrailerRaw);
            }

            if (!previousOffset.HasValue) {
                break;
            }

            currentOffset = previousOffset.Value;
        }

        if (trailers.Count == 0) {
            return false;
        }

        trailerRaw = string.Join("\n", trailers);
        return true;
    }

    private static bool TryGetXrefStreamTrailerRawAtOffset(
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int xrefStreamOffset,
        out string trailerRaw) {
        trailerRaw = string.Empty;
        foreach (var entry in map.Values) {
            if (!parsedOffsets.TryGetValue(entry.ObjectNumber, out int offset) ||
                offset != xrefStreamOffset ||
                entry.Value is not PdfStream stream ||
                stream.Dictionary.Get<PdfName>("Type")?.Name != "XRef") {
                continue;
            }

            trailerRaw = BuildXrefStreamTrailerRaw(stream.Dictionary);
            return true;
        }

        return false;
    }

    internal static void ThrowIfEncrypted(string trailerRaw) {
        if (ContainsPdfName(trailerRaw, "Encrypt")) {
            throw new NotSupportedException("Encrypted PDF files are not supported by OfficeIMO.Pdf yet.");
        }
    }

    internal static void ThrowIfUnsafeForRewrite(byte[] pdf) {
        if (HasSignatureMarkers(pdf)) {
            throw new NotSupportedException("Signed PDF files are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasFormMarkers(pdf)) {
            throw new NotSupportedException("PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedOutlineRewriteMarkers(pdf)) {
            throw new NotSupportedException("PDF outlines are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedPageLabelRewriteMarkers(pdf)) {
            throw new NotSupportedException("PDF page labels are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedCatalogNameTreeRewriteMarkers(pdf)) {
            throw new NotSupportedException("PDF catalog name trees are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedNamedDestinationRewriteMarkers(pdf)) {
            throw new NotSupportedException("PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedOpenActionRewriteMarkers(pdf)) {
            throw new NotSupportedException("PDF open actions are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedViewerPreferenceRewriteMarkers(pdf)) {
            throw new NotSupportedException("PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasTaggedContentMarkers(pdf)) {
            throw new NotSupportedException("PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedXmpMetadataRewriteMarkers(pdf)) {
            throw new NotSupportedException("PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedCatalogUriRewriteMarkers(pdf)) {
            throw new NotSupportedException("PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedOutputIntentRewriteMarkers(pdf)) {
            throw new NotSupportedException("PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedEmbeddedFileRewriteMarkers(pdf)) {
            throw new NotSupportedException("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasUnsupportedOptionalContentRewriteMarkers(pdf)) {
            throw new NotSupportedException("PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.");
        }

        if (HasActiveContentMarkers(pdf)) {
            throw new NotSupportedException("PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
        }
    }

    internal static bool HasEncryptionMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "Encrypt");
    }

    internal static bool HasSignatureMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "ByteRange", "SigFlags", "Sig") ||
            ContainsAnyParsedPdfName(pdf, "ByteRange", "SigFlags", "Sig");
    }

    internal static bool HasFormMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "AcroForm", "Fields", "FT") ||
            ContainsAnyParsedPdfName(pdf, "AcroForm", "Fields", "FT");
    }

    internal static bool HasAnnotationMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "Annots", "Annot") ||
            ContainsAnyParsedPdfName(pdf, "Annots", "Annot");
    }

    internal static bool HasOutlineMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "Outlines", "UseOutlines") ||
            ContainsAnyParsedPdfName(pdf, "Outlines", "UseOutlines");
    }

    internal static bool HasUnsupportedOutlineRewriteMarkers(byte[] pdf) {
        if (!HasOutlineMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("Outlines", out var outlines)) {
                return false;
            }

            return !IsSupportedOutlineGraph(objects, outlines, new HashSet<int>());
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasCatalogViewSettingMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "PageMode", "PageLayout") ||
            ContainsAnyParsedPdfName(pdf, "PageMode", "PageLayout");
    }

    internal static bool HasPageLabelMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "PageLabels") ||
            ContainsAnyParsedPdfName(pdf, "PageLabels");
    }

    internal static bool HasUnsupportedPageLabelRewriteMarkers(byte[] pdf) {
        if (!HasPageLabelMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("PageLabels", out var pageLabels)) {
                return catalog is null;
            }

            return !IsSupportedPageLabelTree(objects, pageLabels);
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasNamedDestinationMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "Dests") ||
            ContainsAnyParsedPdfName(pdf, "Dests");
    }

    internal static bool HasCatalogNameTreeMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "Names") ||
            ContainsAnyParsedPdfName(pdf, "Names");
    }

    internal static bool HasUnsupportedCatalogNameTreeRewriteMarkers(byte[] pdf) {
        if (!HasCatalogNameTreeMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("Names", out var names)) {
                return false;
            }

            PdfDictionary? namesDictionary = ResolveObject(objects, names) as PdfDictionary;
            if (namesDictionary is null) {
                return true;
            }

            foreach (var key in namesDictionary.Items.Keys) {
                if (string.Equals(key, "Dests", StringComparison.Ordinal) ||
                    string.Equals(key, "EmbeddedFiles", StringComparison.Ordinal) ||
                    string.Equals(key, "JavaScript", StringComparison.Ordinal)) {
                    continue;
                }

                return true;
            }

            return false;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasUnsupportedNamedDestinationRewriteMarkers(byte[] pdf) {
        if (!HasNamedDestinationMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null) {
                return true;
            }

            if (catalog.Items.ContainsKey("Dests")) {
                return false;
            }

            if (catalog.Items.TryGetValue("Names", out var names)) {
                PdfDictionary? namesDictionary = ResolveObject(objects, names) as PdfDictionary;
                if (namesDictionary is null) {
                    return true;
                }

                if (namesDictionary.Items.ContainsKey("Dests")) {
                    return !TryGetNamedDestinationNameTree(objects, names, out var namedDestinationTree) ||
                        !IsSupportedNamedDestinationNameTree(objects, namedDestinationTree);
                }
            }

            return false;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasOpenActionMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "OpenAction") ||
            ContainsAnyParsedPdfName(pdf, "OpenAction");
    }

    internal static bool HasUnsupportedOpenActionRewriteMarkers(byte[] pdf) {
        if (!HasOpenActionMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("OpenAction", out var openAction)) {
                return catalog is null;
            }

            PdfObject? resolved = ResolveObject(objects, openAction);
            if (resolved is PdfArray array &&
                IsDestinationForKnownPage(objects, array)) {
                return false;
            }

            if (resolved is PdfDictionary dictionary &&
                IsSupportedGoToActionDictionary(objects, dictionary)) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasViewerPreferenceMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "ViewerPreferences") ||
            ContainsAnyParsedPdfName(pdf, "ViewerPreferences");
    }

    internal static bool HasUnsupportedViewerPreferenceRewriteMarkers(byte[] pdf) {
        if (!HasViewerPreferenceMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("ViewerPreferences", out var viewerPreferences)) {
                return catalog is null;
            }

            PdfObject? resolved = ResolveObject(objects, viewerPreferences);
            if (resolved is PdfDictionary dictionary &&
                IsSimpleCatalogDictionary(dictionary)) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasTaggedContentMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "MarkInfo", "StructTreeRoot", "ParentTree", "StructElem") ||
            ContainsAnyParsedPdfName(pdf, "MarkInfo", "StructTreeRoot", "ParentTree", "StructElem");
    }

    internal static bool HasXmpMetadataMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "Metadata") ||
            ContainsAnyParsedPdfName(pdf, "Metadata");
    }

    internal static bool HasUnsupportedXmpMetadataRewriteMarkers(byte[] pdf) {
        if (!HasXmpMetadataMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("Metadata", out var xmpMetadata)) {
                return catalog is null;
            }

            if (IsSupportedCatalogXmpMetadataStream(objects, xmpMetadata)) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasCatalogUriMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        if (!ContainsPdfName(text, "URI")) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            return catalog?.Items.ContainsKey("URI") == true;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasUnsupportedCatalogUriRewriteMarkers(byte[] pdf) {
        if (!HasCatalogUriMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("URI", out var catalogUri)) {
                return catalog is null;
            }

            PdfObject? resolved = ResolveObject(objects, catalogUri);
            if (resolved is PdfDictionary dictionary &&
                IsSimpleCatalogDictionary(dictionary)) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasOutputIntentMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsPdfName(text, "OutputIntents") ||
            ContainsPdfName(text, "OutputIntent");
    }

    internal static bool HasUnsupportedOutputIntentRewriteMarkers(byte[] pdf) {
        if (!HasOutputIntentMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("OutputIntents", out var outputIntents)) {
                return catalog is null;
            }

            if (IsSupportedCatalogMetadataGraph(objects, outputIntents, new HashSet<int>())) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasEmbeddedFileMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "EmbeddedFiles", "Filespec", "EmbeddedFile", "AF") ||
            ContainsAnyParsedPdfName(pdf, "EmbeddedFiles", "Filespec", "EmbeddedFile", "AF");
    }

    internal static bool HasUnsupportedEmbeddedFileRewriteMarkers(byte[] pdf) {
        if (!HasEmbeddedFileMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null) {
                return true;
            }

            if (catalog.Items.TryGetValue("Names", out var names)) {
                PdfDictionary? namesDictionary = ResolveObject(objects, names) as PdfDictionary;
                if (namesDictionary is null) {
                    return true;
                }

                if (namesDictionary.Items.ContainsKey("EmbeddedFiles")) {
                    if (!TryGetEmbeddedFilesNameTree(objects, names, out var embeddedFiles) ||
                        !IsSupportedCatalogMetadataGraph(objects, embeddedFiles, new HashSet<int>())) {
                        return true;
                    }
                }
            }

            if (catalog.Items.TryGetValue("AF", out var associatedFiles)) {
                if (!IsSupportedCatalogMetadataGraph(objects, associatedFiles, new HashSet<int>())) {
                    return true;
                }
            }

            return false;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasOptionalContentMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "OCProperties", "OCGs", "OCG", "OCMD") ||
            ContainsAnyParsedPdfName(pdf, "OCProperties", "OCGs", "OCG", "OCMD");
    }

    internal static bool HasUnsupportedOptionalContentRewriteMarkers(byte[] pdf) {
        if (!HasOptionalContentMarkers(pdf)) {
            return false;
        }

        try {
            var (objects, trailerRaw) = ParseObjects(pdf);
            PdfDictionary? catalog = FindCatalog(objects, trailerRaw);
            if (catalog is null ||
                !catalog.Items.TryGetValue("OCProperties", out var optionalContent)) {
                return catalog is null;
            }

            if (IsSupportedCatalogMetadataGraph(objects, optionalContent, new HashSet<int>())) {
                return false;
            }

            return true;
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return true;
        }
    }

    internal static bool HasActiveContentMarkers(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        string text = PdfEncoding.Latin1GetString(pdf);
        return ContainsAnyPdfName(text, "JavaScript", "JS", "AA", "Launch", "SubmitForm", "RichMedia") ||
            ContainsAnyParsedPdfName(pdf, "JavaScript", "JS", "AA", "Launch", "SubmitForm", "RichMedia");
    }

    internal static string? GetHeaderVersion(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));

        if (pdf.Length < 8 ||
            pdf[0] != (byte)'%' ||
            pdf[1] != (byte)'P' ||
            pdf[2] != (byte)'D' ||
            pdf[3] != (byte)'F' ||
            pdf[4] != (byte)'-') {
            return null;
        }

        int start = 5;
        int end = start;
        while (end < pdf.Length) {
            byte value = pdf[end];
            if (value == (byte)'\r' || value == (byte)'\n' || value == (byte)' ' || value == (byte)'\t') {
                break;
            }

            end++;
        }

        return end > start ? PdfEncoding.Latin1GetString(pdf, start, end - start) : null;
    }

    private static void ThrowIfEncryptedXrefStream(Dictionary<int, PdfIndirectObject> map) {
        foreach (var entry in map.Values) {
            PdfDictionary? dictionary = entry.Value switch {
                PdfDictionary directDictionary => directDictionary,
                PdfStream stream => stream.Dictionary,
                _ => null
            };

            if (dictionary is not null && dictionary.Items.ContainsKey("Encrypt")) {
                throw new NotSupportedException("Encrypted PDF files are not supported by OfficeIMO.Pdf yet.");
            }
        }
    }

    internal static PdfDictionary? FindCatalog(Dictionary<int, PdfIndirectObject> map, string? trailerRaw = null) {
        if (TryGetTrailerRootReference(trailerRaw, out PdfReference rootReference) &&
            PdfObjectLookup.TryGet(map, rootReference, out var rootObject) &&
            rootObject.Value is PdfDictionary rootDictionary &&
            rootDictionary.Get<PdfName>("Type")?.Name == "Catalog") {
            return rootDictionary;
        }

        if (TryGetXrefStreamRootReference(map, out rootReference) &&
            PdfObjectLookup.TryGet(map, rootReference, out rootObject) &&
            rootObject.Value is PdfDictionary xrefRootDictionary &&
            xrefRootDictionary.Get<PdfName>("Type")?.Name == "Catalog") {
            return xrefRootDictionary;
        }

        return FindCatalogByScan(map);
    }

    private static PdfDictionary? FindCatalogByScan(Dictionary<int, PdfIndirectObject> map) {
        foreach (var entry in map.Values) {
            if (entry.Value is PdfDictionary dictionary &&
                dictionary.Get<PdfName>("Type")?.Name == "Catalog") {
                return dictionary;
            }
        }

        return null;
    }

    private static bool TryGetTrailerRootReference(string? trailerRaw, out PdfReference reference) {
        reference = null!;
        if (string.IsNullOrWhiteSpace(trailerRaw)) {
            return false;
        }

        Match match = TrailerRootRegex.Match(trailerRaw);
        if (!match.Success ||
            !int.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int objectNumber) ||
            !int.TryParse(match.Groups[2].Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int generation)) {
            return false;
        }

        reference = new PdfReference(objectNumber, generation);
        return true;
    }

    private static bool TryGetXrefStreamRootReference(Dictionary<int, PdfIndirectObject> map, out PdfReference reference) {
        reference = null!;
        foreach (var entry in map.Values.OrderByDescending(static item => item.ObjectNumber)) {
            PdfDictionary? dictionary = entry.Value switch {
                PdfStream stream => stream.Dictionary,
                PdfDictionary directDictionary => directDictionary,
                _ => null
            };

            if (dictionary?.Get<PdfName>("Type")?.Name == "XRef" &&
                dictionary.Items.TryGetValue("Root", out var root) &&
                root is PdfReference rootReference &&
                ResolveObject(map, rootReference) is PdfDictionary rootDictionary &&
                rootDictionary.Get<PdfName>("Type")?.Name == "Catalog" &&
                rootReference.Generation >= 0) {
                reference = rootReference;
                return true;
            }
        }

        return false;
    }

    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> map, PdfObject? value) {
        return PdfObjectLookup.Resolve(map, value);
    }

    private static bool IsDestinationForKnownPage(Dictionary<int, PdfIndirectObject> map, PdfArray destination) {
        return destination.Items.Count > 0 &&
            destination.Items[0] is PdfReference pageReference &&
            PdfObjectLookup.TryGet(map, pageReference, out var pageObject) &&
            pageObject.Value is PdfDictionary pageDictionary &&
            pageDictionary.Get<PdfName>("Type")?.Name == "Page";
    }

    private static bool IsDestinationForKnownPage(Dictionary<int, PdfIndirectObject> map, PdfObject destination) {
        var visitedReferences = new HashSet<(int ObjectNumber, int Generation)>();
        while (true) {
            if (destination is PdfReference reference) {
                if (!visitedReferences.Add((reference.ObjectNumber, reference.Generation)) ||
                    !PdfObjectLookup.TryGet(map, reference, out var indirect)) {
                    return false;
                }

                destination = indirect.Value;
                continue;
            }

            if (destination is PdfDictionary dictionary &&
                dictionary.Items.TryGetValue("D", out var explicitDestination)) {
                destination = explicitDestination;
                continue;
            }

            return destination is PdfArray array && IsDestinationForKnownPage(map, array);
        }
    }

    private static bool IsSupportedGoToActionDictionary(Dictionary<int, PdfIndirectObject> map, PdfDictionary dictionary) {
        return dictionary.Items.Count == 2 &&
            dictionary.Get<PdfName>("S")?.Name == "GoTo" &&
            dictionary.Items.TryGetValue("D", out var destination) &&
            IsDestinationForKnownPage(map, destination);
    }

    private static bool IsSupportedOutlineAction(Dictionary<int, PdfIndirectObject> map, PdfObject action) {
        return ResolveObject(map, action) is PdfDictionary dictionary &&
            IsSupportedGoToActionDictionary(map, dictionary);
    }

    private static bool IsSimpleCatalogDictionary(PdfDictionary dictionary) {
        foreach (var value in dictionary.Items.Values) {
            if (!IsSimpleCatalogValue(value)) {
                return false;
            }
        }

        return true;
    }

    private static bool IsSimpleCatalogValue(PdfObject value) {
        switch (value) {
            case PdfNumber:
            case PdfBoolean:
            case PdfName:
            case PdfStringObj:
            case PdfNull:
                return true;
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!IsSimpleCatalogValue(item)) {
                        return false;
                    }
                }

                return true;
            default:
                return false;
        }
    }

    private static bool IsSupportedCatalogXmpMetadataStream(Dictionary<int, PdfIndirectObject> map, PdfObject value) {
        if (value is not PdfReference reference ||
            !PdfObjectLookup.TryGet(map, reference, out var indirect) ||
            indirect.Value is not PdfStream stream ||
            stream.Dictionary.Get<PdfName>("Type")?.Name != "Metadata" ||
            stream.Dictionary.Get<PdfName>("Subtype")?.Name != "XML") {
            return false;
        }

        foreach (var entry in stream.Dictionary.Items) {
            if (string.Equals(entry.Key, "Length", StringComparison.Ordinal)) {
                continue;
            }

            if (!IsSimpleCatalogValue(entry.Value)) {
                return false;
            }
        }

        return true;
    }

    private static bool IsSupportedCatalogMetadataGraph(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject value,
        HashSet<int> visitedReferences) {
        switch (value) {
            case PdfNumber:
            case PdfBoolean:
            case PdfName:
            case PdfStringObj:
            case PdfNull:
                return true;
            case PdfReference reference:
                if (!visitedReferences.Add(reference.ObjectNumber)) {
                    return true;
                }

                if (!PdfObjectLookup.TryGet(map, reference, out var indirect)) {
                    return false;
                }

                return !IsPageDictionary(indirect.Value) &&
                    IsSupportedCatalogMetadataGraph(map, indirect.Value, visitedReferences);
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!IsSupportedCatalogMetadataGraph(map, item, visitedReferences)) {
                        return false;
                    }
                }

                return true;
            case PdfDictionary dictionary:
                if (IsPageDictionary(dictionary)) {
                    return false;
                }

                foreach (var item in dictionary.Items.Values) {
                    if (!IsSupportedCatalogMetadataGraph(map, item, visitedReferences)) {
                        return false;
                    }
                }

                return true;
            case PdfStream stream:
                if (IsPageDictionary(stream.Dictionary)) {
                    return false;
                }

                foreach (var item in stream.Dictionary.Items.Values) {
                    if (!IsSupportedCatalogMetadataGraph(map, item, visitedReferences)) {
                        return false;
                    }
                }

                return true;
            default:
                return false;
        }
    }

    private static bool IsSupportedOutlineGraph(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject value,
        HashSet<int> visitedReferences) {
        switch (value) {
            case PdfNumber:
            case PdfBoolean:
            case PdfName:
            case PdfStringObj:
            case PdfNull:
                return true;
            case PdfReference reference:
                if (!visitedReferences.Add(reference.ObjectNumber)) {
                    return true;
                }

                if (!PdfObjectLookup.TryGet(map, reference, out var indirect)) {
                    return false;
                }

                return IsPageDictionary(indirect.Value) ||
                    IsSupportedOutlineGraph(map, indirect.Value, visitedReferences);
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!IsSupportedOutlineGraph(map, item, visitedReferences)) {
                        return false;
                    }
                }

                return true;
            case PdfDictionary dictionary:
                if (IsPageDictionary(dictionary)) {
                    return true;
                }

                if (dictionary.Items.ContainsKey("AA")) {
                    return false;
                }

                if (dictionary.Items.TryGetValue("A", out var action) &&
                    !IsSupportedOutlineAction(map, action)) {
                    return false;
                }

                foreach (var item in dictionary.Items.Values) {
                    if (!IsSupportedOutlineGraph(map, item, visitedReferences)) {
                        return false;
                    }
                }

                return true;
            default:
                return false;
        }
    }

    private static bool IsPageDictionary(PdfObject value) {
        return value is PdfDictionary dictionary &&
            dictionary.Get<PdfName>("Type")?.Name == "Page";
    }

    private static bool TryGetEmbeddedFilesNameTree(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject names,
        out PdfObject embeddedFiles) {
        embeddedFiles = PdfNull.Instance;
        PdfDictionary? namesDictionary = ResolveObject(map, names) as PdfDictionary;
        if (namesDictionary is null ||
            !namesDictionary.Items.TryGetValue("EmbeddedFiles", out var embeddedFileTree)) {
            return false;
        }

        embeddedFiles = embeddedFileTree;
        return true;
    }

    private static bool TryGetNamedDestinationNameTree(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject names,
        out PdfObject namedDestinations) {
        namedDestinations = PdfNull.Instance;
        PdfDictionary? namesDictionary = ResolveObject(map, names) as PdfDictionary;
        if (namesDictionary is null ||
            !namesDictionary.Items.TryGetValue("Dests", out var namedDestinationTree)) {
            return false;
        }

        namedDestinations = namedDestinationTree;
        return true;
    }

    private static bool IsSupportedPageLabelTree(Dictionary<int, PdfIndirectObject> map, PdfObject pageLabels) {
        PdfDictionary? tree = ResolveObject(map, pageLabels) as PdfDictionary;
        if (tree is null ||
            tree.Items.ContainsKey("Kids") ||
            !tree.Items.TryGetValue("Nums", out var numsObject) ||
            ResolveObject(map, numsObject) is not PdfArray nums ||
            nums.Items.Count % 2 != 0) {
            return false;
        }

        for (int i = 0; i < nums.Items.Count; i += 2) {
            if (ResolveObject(map, nums.Items[i]) is not PdfNumber pageIndex ||
                pageIndex.Value < 0 ||
                pageIndex.Value > int.MaxValue ||
                Math.Truncate(pageIndex.Value) != pageIndex.Value ||
                ResolveObject(map, nums.Items[i + 1]) is not PdfDictionary labelDictionary) {
                return false;
            }

            foreach (var value in labelDictionary.Items.Values) {
                if (!IsSimpleCatalogValue(value)) {
                    return false;
                }
            }
        }

        return true;
    }

    private static bool IsSupportedNamedDestinationNameTree(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject namedDestinations) {
        return TryCollectNamedDestinationNameTreeEntries(map, namedDestinations, new HashSet<int>());
    }

    private static bool TryCollectNamedDestinationNameTreeEntries(
        Dictionary<int, PdfIndirectObject> map,
        PdfObject value,
        HashSet<int> visitedReferences) {
        if (value is PdfReference reference) {
            if (!visitedReferences.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(map, reference, out var indirect)) {
                return false;
            }

            return TryCollectNamedDestinationNameTreeEntries(map, indirect.Value, visitedReferences);
        }

        if (value is not PdfDictionary tree) {
            return false;
        }

        bool hasNames = tree.Items.TryGetValue("Names", out var namesObject);
        bool hasKids = tree.Items.TryGetValue("Kids", out var kidsObject);
        if (hasNames && hasKids) {
            return false;
        }

        if (hasNames) {
            if (ResolveObject(map, namesObject) is not PdfArray names ||
                names.Items.Count % 2 != 0) {
                return false;
            }

            for (int i = 0; i < names.Items.Count; i += 2) {
                if (names.Items[i] is not PdfStringObj) {
                    return false;
                }

                PdfObject? destination = ResolveObject(map, names.Items[i + 1]);
                if (destination is null || !IsDestinationForKnownPage(map, destination)) {
                    return false;
                }
            }
        }

        if (hasKids) {
            if (ResolveObject(map, kidsObject) is not PdfArray kids) {
                return false;
            }

            foreach (var kid in kids.Items) {
                if (kid is not PdfReference) {
                    return false;
                }

                if (!TryCollectNamedDestinationNameTreeEntries(map, kid, visitedReferences)) {
                    return false;
                }
            }
        }

        return hasNames || hasKids;
    }

    private static bool ContainsPdfName(string text, string name) {
        if (string.IsNullOrEmpty(text)) return false;

        string token = "/" + name;
        int index = 0;
        while (index < text.Length) {
            index = text.IndexOf(token, index, StringComparison.Ordinal);
            if (index < 0) return false;

            int after = index + token.Length;
            if (after >= text.Length || IsPdfDelimiter(text[after]) || char.IsWhiteSpace(text[after])) {
                return true;
            }

            index = after;
        }

        return false;
    }

    private static bool ContainsAnyPdfName(string text, params string[] names) {
        for (int i = 0; i < names.Length; i++) {
            if (ContainsPdfName(text, names[i])) {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsAnyParsedPdfName(byte[] pdf, params string[] names) {
        try {
            var (map, _) = ParseObjects(pdf);
            var nameSet = new HashSet<string>(names, StringComparer.Ordinal);
            foreach (PdfIndirectObject indirectObject in map.Values) {
                if (ContainsAnyParsedPdfName(indirectObject.Value, nameSet)) {
                    return true;
                }
            }
        } catch (Exception ex) when (ex is not OutOfMemoryException && ex is not StackOverflowException) {
            return false;
        }

        return false;
    }

    private static bool ContainsAnyParsedPdfName(PdfObject value, HashSet<string> names) {
        switch (value) {
            case PdfName name:
                return names.Contains(name.Name);
            case PdfDictionary dictionary:
                foreach (var item in dictionary.Items) {
                    if (names.Contains(item.Key) || ContainsAnyParsedPdfName(item.Value, names)) {
                        return true;
                    }
                }

                return false;
            case PdfArray array:
                foreach (PdfObject item in array.Items) {
                    if (ContainsAnyParsedPdfName(item, names)) {
                        return true;
                    }
                }

                return false;
            case PdfStream stream:
                return ContainsAnyParsedPdfName(stream.Dictionary, names);
            default:
                return false;
        }
    }

    private static bool IsPdfDelimiter(char value) {
        switch (value) {
            case '(':
            case ')':
            case '<':
            case '>':
            case '[':
            case ']':
            case '{':
            case '}':
            case '/':
            case '%':
                return true;
            default:
                return false;
        }
    }

    private static void ResolveIndirectStreamLengths(Dictionary<int, PdfIndirectObject> map, byte[] pdf, List<(int Id, int Generation, int DataStart)> streamLocations) {
        foreach (var streamLocation in streamLocations) {
            if (!map.TryGetValue(streamLocation.Id, out var indirect) || indirect.Value is not PdfStream stream) {
                continue;
            }

            if (!TryGetResolvedLength(stream.Dictionary, map, out int byteLen)) {
                continue;
            }

            int byteStart = streamLocation.DataStart;
            if (byteStart < 0 || byteLen < 0 || byteStart + byteLen > pdf.Length) {
                continue;
            }

            if (stream.Data.Length == byteLen) {
                continue;
            }

            var data = new byte[byteLen];
            Buffer.BlockCopy(pdf, byteStart, data, 0, byteLen);
            map[streamLocation.Id] = new PdfIndirectObject(streamLocation.Id, streamLocation.Generation, new PdfStream(stream.Dictionary, data, stream.DecodingFailed, stream.DecodingError));
        }
    }

    private static bool ApplyClassicXrefEntries(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        HashSet<int> activeObjectNumbers,
        out bool appliedClassicEntries) {
        appliedClassicEntries = false;
        string text = PdfEncoding.Latin1GetString(pdf);
        if (!TryGetLatestStartXrefOffset(text, out int activeXrefOffset)) {
            return false;
        }

        var tables = GetClassicXrefTableChain(text, activeXrefOffset);
        if (tables.Count == 0) {
            return false;
        }

        appliedClassicEntries = true;
        bool appliedXrefStream = false;
        foreach (var table in tables) {
            ApplyClassicXrefTableEntries(map, pdf, parsedOffsets, text, table.Entries, activeObjectNumbers);
            if (table.XrefStreamOffset.HasValue) {
                appliedXrefStream = ApplyXrefStreamAtOffset(map, pdf, parsedOffsets, text, table.XrefStreamOffset.Value) || appliedXrefStream;
            }
        }

        return appliedXrefStream;
    }

    private static void ApplyClassicXrefTableEntries(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        string text,
        List<(int ObjectNumber, int Offset, int Generation, bool InUse)> entries,
        HashSet<int>? activeObjectNumbers = null) {
        foreach (var entry in entries) {
            if (!entry.InUse) {
                if (entry.ObjectNumber != 0) {
                    map.Remove(entry.ObjectNumber);
                    parsedOffsets.Remove(entry.ObjectNumber);
                    activeObjectNumbers?.Remove(entry.ObjectNumber);
                }

                continue;
            }

            if (entry.Offset <= 0 ||
                entry.Offset >= pdf.Length) {
                continue;
            }

            if (TryParseIndirectObjectAt(pdf, text, entry.Offset, map, out var parsed) &&
                parsed.ObjectNumber == entry.ObjectNumber &&
                parsed.Generation == entry.Generation) {
                map[entry.ObjectNumber] = parsed;
                parsedOffsets[entry.ObjectNumber] = entry.Offset;
                activeObjectNumbers?.Add(entry.ObjectNumber);
            }
        }
    }

    private static List<(int Offset, List<(int ObjectNumber, int Offset, int Generation, bool InUse)> Entries, int? XrefStreamOffset)> GetClassicXrefTableChain(string text, int activeXrefOffset) {
        var newestToOldest = new List<(int Offset, List<(int ObjectNumber, int Offset, int Generation, bool InUse)> Entries, int? XrefStreamOffset)>();
        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (visited.Add(currentOffset) &&
            newestToOldest.Count < 64 &&
            TryParseClassicXrefTable(text, currentOffset, out var entries, out int? previousOffset, out _, out int? xrefStreamOffset)) {
            newestToOldest.Add((currentOffset, entries, xrefStreamOffset));
            if (!previousOffset.HasValue) {
                break;
            }

            currentOffset = previousOffset.Value;
        }

        newestToOldest.Reverse();
        return newestToOldest;
    }

    private static bool TryParseClassicXrefTable(string text, int offset, out List<(int ObjectNumber, int Offset, int Generation, bool InUse)> entries, out int? previousOffset, out string trailerRaw, out int? xrefStreamOffset) {
        entries = new List<(int ObjectNumber, int Offset, int Generation, bool InUse)>();
        previousOffset = null;
        trailerRaw = string.Empty;
        xrefStreamOffset = null;
        if (offset < 0 ||
            offset + 4 > text.Length ||
            !string.Equals(text.Substring(offset, 4), "xref", StringComparison.Ordinal) ||
            !HasKeywordBoundary(text, offset - 1, 0, text.Length) ||
            !HasKeywordBoundary(text, offset + 4, 0, text.Length)) {
            return false;
        }

        int trailerIndex = IndexOfKeyword(text, "trailer", offset + 4, text.Length);
        if (trailerIndex < 0) {
            return false;
        }

        string section = SafeSlice(text, offset + 4, trailerIndex - (offset + 4), 2_000_000);
        using (var reader = new StringReader(section)) {
            string? line;
            while ((line = reader.ReadLine()) is not null) {
                string[] headerParts = SplitWhitespace(line);
                if (headerParts.Length < 2 ||
                    !int.TryParse(headerParts[0], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int firstObjectNumber) ||
                    !int.TryParse(headerParts[1], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int count) ||
                    firstObjectNumber < 0 ||
                    count <= 0 ||
                    count > 1_000_000) {
                    continue;
                }

                for (int i = 0; i < count; i++) {
                    string? entryLine = reader.ReadLine();
                    if (entryLine is null) {
                        return entries.Count > 0;
                    }

                    string[] entryParts = SplitWhitespace(entryLine);
                    if (entryParts.Length < 3 ||
                        !int.TryParse(entryParts[0], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int objectOffset) ||
                        !int.TryParse(entryParts[1], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int generation)) {
                        continue;
                    }

                    if (string.Equals(entryParts[2], "n", StringComparison.Ordinal)) {
                        entries.Add((firstObjectNumber + i, objectOffset, generation, true));
                    } else if (string.Equals(entryParts[2], "f", StringComparison.Ordinal)) {
                        entries.Add((firstObjectNumber + i, objectOffset, generation, false));
                    }
                }
            }
        }

        if (entries.Count == 0) {
            return false;
        }

        int dictStart = text.IndexOf("<<", trailerIndex, StringComparison.Ordinal);
        if (dictStart >= 0) {
            int dictEnd = FindDictEnd(text, dictStart, text.Length);
            if (dictEnd > dictStart) {
                trailerRaw = SafeSlice(text, trailerIndex, dictEnd - trailerIndex, 1_000_000);
                string dictText = SafeSlice(text, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000);
                try {
                    PdfDictionary trailer = ParseDictionary(dictText);
                    if (trailer.Get<PdfNumber>("Prev") is PdfNumber previous &&
                        previous.Value >= 0 &&
                        previous.Value <= int.MaxValue) {
                        previousOffset = (int)Math.Floor(previous.Value);
                    }

                    if (trailer.Get<PdfNumber>("XRefStm") is PdfNumber xrefStream &&
                        xrefStream.Value >= 0 &&
                        xrefStream.Value <= int.MaxValue) {
                        xrefStreamOffset = (int)Math.Floor(xrefStream.Value);
                    }
                } catch (Exception ex) when (ex is not OutOfMemoryException) {
                    previousOffset = null;
                    xrefStreamOffset = null;
                }
            }
        }

        return true;
    }

    private static string[] SplitWhitespace(string value) {
        return value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
    }

    private static bool ApplyXrefStreamEntries(Dictionary<int, PdfIndirectObject> map, byte[] pdf, Dictionary<int, int> parsedOffsets) {
        var xrefStreams = new List<(int ObjectNumber, int Offset, PdfStream Stream)>();
        foreach (var entry in map.Values) {
            if (entry.Value is PdfStream stream &&
                stream.Dictionary.Get<PdfName>("Type")?.Name == "XRef") {
                int offset = parsedOffsets.TryGetValue(entry.ObjectNumber, out int parsedOffset) ? parsedOffset : int.MaxValue;
                xrefStreams.Add((entry.ObjectNumber, offset, stream));
            }
        }

        if (xrefStreams.Count == 0) {
            return false;
        }

        string text = PdfEncoding.Latin1GetString(pdf);
        if (!TryGetLatestStartXrefOffset(text, out int activeXrefOffset)) {
            return false;
        }

        xrefStreams.Sort(static (left, right) => left.Offset.CompareTo(right.Offset));
        var activeChainOffsets = GetXrefStreamChainOffsets(xrefStreams, activeXrefOffset);
        if (activeChainOffsets.Count == 0) {
            return false;
        }

        var classicPredecessors = GetClassicPredecessorTablesForXrefStreamChain(text, xrefStreams, activeXrefOffset);
        foreach (var table in classicPredecessors) {
            ApplyClassicXrefTableEntries(map, pdf, parsedOffsets, text, table.Entries);
            if (table.XrefStreamOffset.HasValue) {
                ApplyXrefStreamAtOffset(map, pdf, parsedOffsets, text, table.XrefStreamOffset.Value);
            }
        }

        foreach (int chainOffset in activeChainOffsets) {
            var xrefStream = xrefStreams.First(item => item.Offset == chainOffset);
            ApplyXrefStreamObjectEntries(map, pdf, parsedOffsets, text, xrefStream.Stream);
        }

        return true;
    }

    private static bool ApplyXrefStreamAtOffset(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        string text,
        int xrefStreamOffset) {
        PdfStream? targetStream = null;
        foreach (var entry in map.Values) {
            if (!parsedOffsets.TryGetValue(entry.ObjectNumber, out int offset) ||
                offset != xrefStreamOffset ||
                entry.Value is not PdfStream stream ||
                stream.Dictionary.Get<PdfName>("Type")?.Name != "XRef") {
                continue;
            }

            targetStream = stream;
            break;
        }

        if (targetStream is null) {
            return false;
        }

        ApplyXrefStreamObjectEntries(map, pdf, parsedOffsets, text, targetStream);
        return true;
    }

    private static void ApplyXrefStreamObjectEntries(
        Dictionary<int, PdfIndirectObject> map,
        byte[] pdf,
        Dictionary<int, int> parsedOffsets,
        string text,
        PdfStream xrefStream) {
        byte[] data = Filters.StreamDecoder.Decode(xrefStream.Dictionary, xrefStream.Data, map);
        var entries = ReadXrefStreamEntries(xrefStream.Dictionary, data).ToList();
        foreach (var entry in entries) {
            if (entry.Type == 0 &&
                entry.ObjectNumber != 0) {
                map.Remove(entry.ObjectNumber);
                parsedOffsets.Remove(entry.ObjectNumber);
            }
        }

        foreach (var entry in entries) {
            if (entry.Type != 1 ||
                entry.Field1 < 0 ||
                entry.Field1 > int.MaxValue ||
                entry.Field2 < 0 ||
                entry.Field2 > int.MaxValue) {
                continue;
            }

            int offset = (int)entry.Field1;
            int generation = (int)entry.Field2;
            if (TryParseIndirectObjectAt(pdf, text, offset, map, out var parsed) &&
                parsed.ObjectNumber == entry.ObjectNumber &&
                parsed.Generation == generation) {
                map[entry.ObjectNumber] = parsed;
                parsedOffsets[entry.ObjectNumber] = offset;
            }
        }

        foreach (var entry in entries) {
            if (entry.Type != 2 ||
                entry.Field1 < 0 ||
                entry.Field1 > int.MaxValue ||
                entry.Field2 < 0 ||
                entry.Field2 > int.MaxValue) {
                continue;
            }

            int objectStreamNumber = (int)entry.Field1;
            int objectStreamIndex = (int)entry.Field2;
            if (TryParseObjectFromObjectStream(map, parsedOffsets, objectStreamNumber, objectStreamIndex, entry.ObjectNumber, out var parsed, out int objectStreamOffset)) {
                map[entry.ObjectNumber] = parsed;
                parsedOffsets[entry.ObjectNumber] = objectStreamOffset;
            }
        }
    }

    private static List<(int Offset, List<(int ObjectNumber, int Offset, int Generation, bool InUse)> Entries, int? XrefStreamOffset)> GetClassicPredecessorTablesForXrefStreamChain(
        string text,
        List<(int ObjectNumber, int Offset, PdfStream Stream)> xrefStreams,
        int activeXrefOffset) {
        var byOffset = new Dictionary<int, PdfStream>();
        foreach (var xrefStream in xrefStreams) {
            byOffset[xrefStream.Offset] = xrefStream.Stream;
        }

        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (byOffset.TryGetValue(currentOffset, out PdfStream? stream) &&
            visited.Add(currentOffset) &&
            visited.Count < 64) {
            if (stream.Dictionary.Get<PdfNumber>("Prev") is not PdfNumber previous ||
                previous.Value < 0 ||
                previous.Value > int.MaxValue) {
                return new List<(int Offset, List<(int ObjectNumber, int Offset, int Generation, bool InUse)> Entries, int? XrefStreamOffset)>();
            }

            currentOffset = (int)Math.Floor(previous.Value);
        }

        return GetClassicXrefTableChain(text, currentOffset);
    }

    private static List<int> GetXrefStreamChainOffsets(List<(int ObjectNumber, int Offset, PdfStream Stream)> xrefStreams, int activeXrefOffset) {
        var byOffset = new Dictionary<int, PdfStream>();
        foreach (var xrefStream in xrefStreams) {
            byOffset[xrefStream.Offset] = xrefStream.Stream;
        }

        var newestToOldest = new List<int>();
        var visited = new HashSet<int>();
        int currentOffset = activeXrefOffset;
        while (byOffset.TryGetValue(currentOffset, out PdfStream? stream) &&
            visited.Add(currentOffset) &&
            newestToOldest.Count < 64) {
            newestToOldest.Add(currentOffset);
            if (stream.Dictionary.Get<PdfNumber>("Prev") is not PdfNumber previous ||
                previous.Value < 0 ||
                previous.Value > int.MaxValue) {
                break;
            }

            currentOffset = (int)Math.Floor(previous.Value);
        }

        newestToOldest.Reverse();
        return newestToOldest;
    }

    private static bool TryGetLatestStartXrefOffset(string text, out int offset) {
        offset = 0;
        int startXrefIndex = text.LastIndexOf("startxref", StringComparison.Ordinal);
        if (startXrefIndex < 0) {
            return false;
        }

        int index = startXrefIndex + "startxref".Length;
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            index++;
        }

        long value = 0;
        int firstDigit = index;
        while (index < text.Length && char.IsDigit(text[index])) {
            value = (value * 10) + (text[index] - '0');
            if (value > int.MaxValue) {
                return false;
            }

            index++;
        }

        if (index == firstDigit) {
            return false;
        }

        offset = (int)value;
        return true;
    }

    private static IEnumerable<XrefStreamEntry> ReadXrefStreamEntries(PdfDictionary dictionary, byte[] data) {
        if (data.Length == 0 ||
            dictionary.Get<PdfArray>("W") is not PdfArray widthsArray ||
            widthsArray.Items.Count < 3) {
            yield break;
        }

        int w0 = GetNonNegativeInt(widthsArray.Items[0]);
        int w1 = GetNonNegativeInt(widthsArray.Items[1]);
        int w2 = GetNonNegativeInt(widthsArray.Items[2]);
        int entryWidth = w0 + w1 + w2;
        if (entryWidth <= 0) {
            yield break;
        }

        var ranges = GetXrefIndexRanges(dictionary);
        int dataOffset = 0;
        foreach (var range in ranges) {
            for (int i = 0; i < range.Count; i++) {
                if (dataOffset + entryWidth > data.Length) {
                    yield break;
                }

                long type = w0 == 0 ? 1 : ReadBigEndian(data, dataOffset, w0);
                dataOffset += w0;
                long field1 = ReadBigEndian(data, dataOffset, w1);
                dataOffset += w1;
                long field2 = ReadBigEndian(data, dataOffset, w2);
                dataOffset += w2;

                yield return new XrefStreamEntry(range.FirstObjectNumber + i, type, field1, field2);
            }
        }
    }

    private static List<(int FirstObjectNumber, int Count)> GetXrefIndexRanges(PdfDictionary dictionary) {
        var ranges = new List<(int, int)>();
        if (dictionary.Get<PdfArray>("Index") is PdfArray indexArray && indexArray.Items.Count >= 2) {
            for (int i = 0; i + 1 < indexArray.Items.Count; i += 2) {
                int first = GetNonNegativeInt(indexArray.Items[i]);
                int count = GetNonNegativeInt(indexArray.Items[i + 1]);
                if (count > 0) {
                    ranges.Add((first, count));
                }
            }
        }

        if (ranges.Count == 0) {
            int size = GetNonNegativeInt(dictionary.Get<PdfNumber>("Size"));
            if (size > 0) {
                ranges.Add((0, size));
            }
        }

        return ranges;
    }

    private static int GetNonNegativeInt(PdfObject? value) {
        if (value is not PdfNumber number || number.Value <= 0) {
            return 0;
        }

        return (int)Math.Min(int.MaxValue, Math.Floor(number.Value));
    }

    private static long ReadBigEndian(byte[] data, int offset, int length) {
        long value = 0;
        for (int i = 0; i < length; i++) {
            value = (value << 8) | data[offset + i];
        }

        return value;
    }

    private static bool TryParseIndirectObjectAt(byte[] pdf, string text, int offset, Dictionary<int, PdfIndirectObject> map, out PdfIndirectObject parsed) {
        parsed = null!;
        if (offset < 0 || offset >= text.Length) {
            return false;
        }

        Match match = ObjRegex.Match(text, offset);
        if (!match.Success || match.Index != offset) {
            return false;
        }

        int id = int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        int gen = int.Parse(match.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
        int start = match.Index;
        int bodyStart = match.Index + match.Length;
        int end = FindObjectEnd(text, start);
        if (end < 0) {
            return false;
        }

        int dictStart = text.IndexOf("<<", start, end - start, System.StringComparison.Ordinal);
        if (dictStart >= 0) {
            int dictEnd = FindDictEnd(text, dictStart, end);
            if (dictEnd > dictStart) {
                string dictText = SafeSlice(text, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000);
                PdfDictionary? dict;
                try { dict = ParseDictionary(dictText); }
                catch (Exception ex) when (ex is not OutOfMemoryException) { dict = null; }
                if (dict is null) {
                    return false;
                }

                int streamKw = IndexOfKeyword(text, "stream", dictEnd, end);
                if (streamKw >= 0) {
                    int dataStart = SkipEOL(text, streamKw + 6, end);
                    int byteLen = -1;
                    TryGetResolvedLength(dict, map, out byteLen);
                    if (byteLen < 0) {
                        int endStream = IndexOfKeyword(text, "endstream", dataStart, end);
                        if (endStream > dataStart) byteLen = endStream - dataStart;
                    }

                    if (byteLen >= 0 && dataStart >= 0 && dataStart + byteLen <= pdf.Length) {
                        var data = new byte[byteLen];
                        Buffer.BlockCopy(pdf, dataStart, data, 0, byteLen);
                        parsed = new PdfIndirectObject(id, gen, new PdfStream(dict, data));
                        return true;
                    }
                }

                parsed = new PdfIndirectObject(id, gen, dict);
                return true;
            }
        }

        int bodyEnd = end;
        if (bodyEnd - 6 >= bodyStart && string.Equals(text.Substring(bodyEnd - 6, 6), "endobj", StringComparison.Ordinal)) {
            bodyEnd -= 6;
        }

        string body = SafeSlice(text, bodyStart, bodyEnd - bodyStart, 1_000_000).Trim();
        var topLevelObject = ParseTopLevelObject(body);
        if (topLevelObject is null) {
            return false;
        }

        parsed = new PdfIndirectObject(id, gen, topLevelObject);
        return true;
    }

    private static bool TryParseObjectFromObjectStream(
        Dictionary<int, PdfIndirectObject> map,
        Dictionary<int, int> parsedOffsets,
        int objectStreamNumber,
        int objectStreamIndex,
        int expectedObjectNumber,
        out PdfIndirectObject parsed,
        out int objectStreamOffset) {
        parsed = null!;
        objectStreamOffset = int.MaxValue;
        if (!map.TryGetValue(objectStreamNumber, out var objectStreamIndirect) ||
            objectStreamIndirect.Value is not PdfStream objectStream ||
            objectStream.Dictionary.Get<PdfName>("Type")?.Name != "ObjStm") {
            return false;
        }

        byte[] data = Filters.StreamDecoder.Decode(objectStream.Dictionary, objectStream.Data, map);
        int n = (int)(objectStream.Dictionary.Get<PdfNumber>("N")?.Value ?? 0);
        int first = (int)(objectStream.Dictionary.Get<PdfNumber>("First")?.Value ?? 0);
        if (objectStreamIndex < 0 || objectStreamIndex >= n || n <= 0 || first <= 0 || first > data.Length) {
            return false;
        }

        var headerBytes = new byte[first];
        Buffer.BlockCopy(data, 0, headerBytes, 0, first);
        string header = PdfEncoding.Latin1GetString(headerBytes);
        var pairs = ParsePairs(header, n);
        if (pairs.Count != n ||
            pairs[objectStreamIndex].Obj != expectedObjectNumber) {
            return false;
        }

        int start = first + pairs[objectStreamIndex].Off;
        int end = (objectStreamIndex + 1 < n) ? first + pairs[objectStreamIndex + 1].Off : data.Length;
        if (start < 0 || end > data.Length || end <= start) {
            return false;
        }

        int len = end - start;
        var sliceBytes = new byte[len];
        Buffer.BlockCopy(data, start, sliceBytes, 0, len);
        var slice = PdfEncoding.Latin1GetString(sliceBytes);
        var parsedObject = ParseTopLevelObject(slice);
        if (parsedObject is null) {
            return false;
        }

        parsed = new PdfIndirectObject(expectedObjectNumber, 0, parsedObject);
        objectStreamOffset = parsedOffsets.TryGetValue(objectStreamNumber, out int offset) ? offset : int.MaxValue;
        return true;
    }

    private static void ExpandObjectStreams(Dictionary<int, PdfIndirectObject> map, byte[] pdf, Dictionary<int, int> parsedOffsets, HashSet<int>? allowedObjectStreamNumbers) {
        // Snapshot keys to avoid modifying during enumeration
        var keys = new List<int>(map.Keys);
        keys.Sort((left, right) => GetSourceOffset(left).CompareTo(GetSourceOffset(right)));
        var effectiveOffsets = new Dictionary<int, int>(parsedOffsets);
        foreach (var id in keys) {
            if (allowedObjectStreamNumbers is not null &&
                !allowedObjectStreamNumbers.Contains(id)) {
                continue;
            }

            if (!map.TryGetValue(id, out var ind)) continue;
            if (ind.Value is not PdfStream s) continue;
            var type = s.Dictionary.Get<PdfName>("Type")?.Name;
            if (!string.Equals(type, "ObjStm", StringComparison.Ordinal)) continue;
            int objectStreamOffset = GetSourceOffset(id);

            // Decode object stream bytes (flate only for now)
            var data = Filters.StreamDecoder.Decode(s.Dictionary, s.Data, map);
            int n = (int)(s.Dictionary.Get<PdfNumber>("N")?.Value ?? 0);
            int first = (int)(s.Dictionary.Get<PdfNumber>("First")?.Value ?? 0);
            if (n <= 0 || first <= 0 || first > data.Length) continue;
            // Header: pairs of objectNumber and offset (ASCII)
            var headerBytes = new byte[first];
            Buffer.BlockCopy(data, 0, headerBytes, 0, first);
            string header = PdfEncoding.Latin1GetString(headerBytes);
            var pairs = ParsePairs(header, n);
            if (pairs.Count != n) continue;
            for (int i = 0; i < n; i++) {
                int objNum = pairs[i].Obj;
                int off = pairs[i].Off;
                if (map.ContainsKey(objNum) &&
                    effectiveOffsets.TryGetValue(objNum, out int currentOffset) &&
                    currentOffset > objectStreamOffset) {
                    continue;
                }

                int start = first + off;
                int end = (i + 1 < n) ? first + pairs[i + 1].Off : data.Length;
                if (start < 0 || end > data.Length || end <= start) continue;
                int len = end - start;
                var sliceBytes = new byte[len];
                Buffer.BlockCopy(data, start, sliceBytes, 0, len);
                var slice = PdfEncoding.Latin1GetString(sliceBytes);
                var parsed = ParseTopLevelObject(slice);
                if (parsed is not null) {
                    map[objNum] = new PdfIndirectObject(objNum, 0, parsed);
                    effectiveOffsets[objNum] = objectStreamOffset;
                }
            }
        }

        int GetSourceOffset(int objectNumber) => parsedOffsets.TryGetValue(objectNumber, out int offset) ? offset : int.MaxValue;
    }

    private static List<(int Obj, int Off)> ParsePairs(string header, int n) {
        var list = new List<(int, int)>(n);
        int i = 0; int count = 0;
        while (i < header.Length && count < n) {
            SkipWs();
            if (!ReadInt(out int obj)) break;
            SkipWs();
            if (!ReadInt(out int off)) break;
            list.Add((obj, off)); count++;
        }
        return list;

        void SkipWs() { while (i < header.Length && char.IsWhiteSpace(header[i])) i++; }
        bool ReadInt(out int val) {
            int sign = 1; if (i < header.Length && header[i] == '-') { sign = -1; i++; }
            int start = i; long v = 0; bool any = false;
            while (i < header.Length && char.IsDigit(header[i])) { v = v * 10 + (header[i] - '0'); i++; any = true; if (i - start > 10) break; }
            val = any ? (int)(v * sign) : 0; return any;
        }
    }

    private readonly struct XrefStreamEntry {
        public XrefStreamEntry(int objectNumber, long type, long field1, long field2) {
            ObjectNumber = objectNumber;
            Type = type;
            Field1 = field1;
            Field2 = field2;
        }

        public int ObjectNumber { get; }
        public long Type { get; }
        public long Field1 { get; }
        public long Field2 { get; }
    }

    private static PdfObject? ParseTopLevelObject(string body) {
        if (string.IsNullOrWhiteSpace(body)) return null;
        var s = body.TrimStart();
        if (string.Equals(s, "true", StringComparison.Ordinal)) return new PdfBoolean(true);
        if (string.Equals(s, "false", StringComparison.Ordinal)) return new PdfBoolean(false);
        if (string.Equals(s, "null", StringComparison.Ordinal)) return PdfNull.Instance;
        if (s.StartsWith("<<", System.StringComparison.Ordinal)) {
            // Find matching >> and parse inside
            int dictStart = body.IndexOf("<<", StringComparison.Ordinal);
            if (dictStart >= 0) {
                int dictEnd = FindDictEnd(body, dictStart, body.Length);
                if (dictEnd > dictStart) {
                    string dictText = SafeSlice(body, dictStart + 2, dictEnd - (dictStart + 2), 1_000_000);
                    try { return ParseDictionary(dictText); } catch { return null; }
                }
            }
            return null;
        }
        if (s.Length > 0 && s[0] == '[') {
            var toks = Tokenize(s);
            var (obj, _) = ParseObject(toks, 0);
            return obj;
        }
        if (s.Length > 0 && s[0] == '(') {
            // literal string
            int end = s.LastIndexOf(')');
            string inner = end > 1 ? s.Substring(1, end - 1) : s.Substring(1);
            return CreateParsedString(PdfTextString.DecodeLiteral(inner));
        }
        if (s.Length > 0 && s[0] == '<' && (s.Length == 1 || s[1] != '<')) {
            int end = s.IndexOf('>');
            string inner = end > 1 ? s.Substring(1, end - 1) : s.Substring(1);
            return CreateParsedString(DecodeHexString(inner));
        }
        // number or name fallbacks
        var tokens = Tokenize(s);
        if (tokens.Count > 0) {
            var (obj0, _) = ParseObject(tokens, 0);
            return obj0;
        }
        return null;
    }

    internal static bool HasFlateDecode(PdfDictionary dict) {
        if (!dict.Items.TryGetValue("Filter", out var f)) return false;
        if (f is PdfName n) return string.Equals(n.Name, "FlateDecode", System.StringComparison.Ordinal);
        if (f is PdfArray arr) {
            foreach (var item in arr.Items) if (item is PdfName nn && string.Equals(nn.Name, "FlateDecode", System.StringComparison.Ordinal)) return true;
        }
        return false;
    }

    private static PdfDictionary ParseDictionary(string dict) {
        var d = new PdfDictionary();
        var tokens = Tokenize(dict);
        for (int i = 0; i < tokens.Count; i++) {
            if (tokens[i].Length > 0 && tokens[i][0] == '/') {
                string key = DecodeName(tokens[i].Substring(1));
                if (i + 1 < tokens.Count) {
                    var (obj, consumed) = ParseObject(tokens, i + 1);
                    d.Items[key] = obj;
                    i += consumed + 1;
                }
            }
        }
        return d;
    }

    private static (PdfObject Obj, int Consumed) ParseObject(List<string> tokens, int i) {
        if (i < 0 || i >= tokens.Count) return (new PdfName(""), 0);
        string tok = tokens[i] ?? string.Empty;
        if (tok == "<<") {
            var dict = new PdfDictionary();
            int j = i + 1;
            while (j < tokens.Count && tokens[j] != ">>") {
                string keyToken = tokens[j] ?? string.Empty;
                if (keyToken.Length > 0 && keyToken[0] == '/') {
                    string key = DecodeName(keyToken.Substring(1));
                    if (j + 1 < tokens.Count) {
                        var (obj, consumed) = ParseObject(tokens, j + 1);
                        dict.Items[key] = obj;
                        j += consumed + 2;
                        continue;
                    }
                }
                j++;
            }
            return (dict, j - i);
        }
        if (tok == "[") {
            var arr = new PdfArray(); int j = i + 1;
            while (j < tokens.Count && tokens[j] != "]") {
                var (inner, used) = ParseObject(tokens, j);
                arr.Items.Add(inner);
                j += used + 1;
            }
            return (arr, j - i);
        }
        if (tok.Length > 0 && tok[0] == '/') return (new PdfName(DecodeName(tok.Substring(1))), 0);
        if (tok.Length > 0 && tok[0] == '(') return (CreateParsedString(PdfTextString.DecodeLiteral(tok.Substring(1, tok.Length - 2))), 0);
        if (tok.Length > 1 && tok[0] == '<' && tok[tok.Length - 1] == '>' && (tok.Length == 2 || tok[1] != '<')) {
            return (CreateParsedString(DecodeHexString(tok.Substring(1, tok.Length - 2))), 0);
        }
        if (string.Equals(tok, "true", StringComparison.Ordinal)) return (new PdfBoolean(true), 0);
        if (string.Equals(tok, "false", StringComparison.Ordinal)) return (new PdfBoolean(false), 0);
        if (string.Equals(tok, "null", StringComparison.Ordinal)) return (PdfNull.Instance, 0);
        if (tok.Length > 0 && (char.IsDigit(tok[0]) || tok[0] == '-' || tok[0] == '+')) {
            // reference (obj gen R) or number
            if (i + 2 < tokens.Count && tokens[i + 2] == "R" && int.TryParse(tokens[i], out int obj) && int.TryParse(tokens[i + 1], out int gen)) {
                return (new PdfReference(obj, gen), 2);
            }
            if (double.TryParse(tok, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val)) {
                return (new PdfNumber(val), 0);
            }
        }
        return (new PdfName(tok), 0);
    }

    private static List<string> Tokenize(string s) {
        // Guardrails for pathological inputs; dictionaries should be small.
        if (s.Length > 1_000_000) s = s.Substring(0, 1_000_000);
        var tokens = new List<string>(Math.Min(16384, s.Length / 2 + 8));
        int i = 0;
        while (i < s.Length) {
            char c = s[i];
            if (char.IsWhiteSpace(c)) { i++; continue; }
            if (c == '%') {
                i++;
                while (i < s.Length && s[i] != '\n' && s[i] != '\r') i++;
                continue;
            }
            if (c == '<' && i + 1 < s.Length && s[i + 1] == '<') { tokens.Add("<<"); i += 2; continue; }
            if (c == '>' && i + 1 < s.Length && s[i + 1] == '>') { tokens.Add(">>"); i += 2; continue; }
            if (c == '[' || c == ']') { tokens.Add(c.ToString()); i++; continue; }
            if (c == '<') {
                int start = i++;
                while (i < s.Length && s[i] != '>') i++;
                if (i < s.Length && s[i] == '>') i++;
                tokens.Add(s.Substring(start, i - start));
                continue;
            }
            if (c == '(') {
                int start = i; i++;
                int depth = 1; bool esc = false;
                var sb = new StringBuilder();
                while (i < s.Length && depth > 0) {
                    char ch = s[i++];
                    if (esc) { sb.Append(ch); esc = false; } else if (ch == '\\') { sb.Append(ch); esc = true; }
                    else if (ch == '(') { depth++; sb.Append(ch); } else if (ch == ')') { depth--; if (depth > 0) sb.Append(ch); } else sb.Append(ch);
                }
                tokens.Add("(" + sb.ToString() + ")");
                continue;
            }
            // name, number, keyword
            int j = i;
            while (j < s.Length && !char.IsWhiteSpace(s[j]) && s[j] != '%' && s[j] != '/' && s[j] != '[' && s[j] != ']' && s[j] != '<' && s[j] != '>' && s[j] != '(' && s[j] != ')') j++;
            string tok = s.Substring(i, j - i);
            if (tok.Length == 0 && s[i] == '/') { // name starting here
                j = i + 1; while (j < s.Length && !char.IsWhiteSpace(s[j]) && s[j] != '%' && s[j] != '/' && s[j] != '[' && s[j] != ']' && s[j] != '<' && s[j] != '>' && s[j] != '(' && s[j] != ')') j++;
                tok = s.Substring(i, j - i);
            }
            tokens.Add(tok);
            if (tokens.Count > 100_000) break; // hard stop
            i = j;
        }
        return tokens;
    }

    private static bool TryGetResolvedLength(PdfDictionary dict, Dictionary<int, PdfIndirectObject> map, out int length) {
        length = -1;

        if (dict.Get<PdfNumber>("Length") is PdfNumber lenNum) {
            int resolved = (int)Math.Max(0, Math.Min(int.MaxValue, lenNum.Value));
            length = resolved;
            return true;
        }

        if (dict.Get<PdfReference>("Length") is PdfReference lenRef &&
            map.TryGetValue(lenRef.ObjectNumber, out var indirectLength) &&
            indirectLength.Value is PdfNumber referencedLength) {
            int resolved = (int)Math.Max(0, Math.Min(int.MaxValue, referencedLength.Value));
            length = resolved;
            return true;
        }

        return false;
    }

    internal static string DecodeName(string raw) {
        if (string.IsNullOrEmpty(raw) || raw.IndexOf('#') < 0) {
            return raw;
        }

        var sb = new StringBuilder(raw.Length);
        for (int i = 0; i < raw.Length; i++) {
            char ch = raw[i];
            if (ch == '#' && i + 2 < raw.Length && TryHexNibble(raw[i + 1], out int hi) && TryHexNibble(raw[i + 2], out int lo)) {
                sb.Append(PdfEncoding.Latin1GetString(new[] { (byte)((hi << 4) | lo) }));
                i += 2;
                continue;
            }

            sb.Append(ch);
        }

        return sb.ToString();
    }

    private static string DecodeHexString(string raw) {
        return PdfTextString.DecodeHex(raw);
    }

    private static PdfStringObj CreateParsedString(string value) {
        return new PdfStringObj(value, useTextStringEncoding: !PdfWinAnsiEncoding.CanEncode(value, out _));
    }

    private static bool TryHexNibble(char c, out int value) {
        if (c >= '0' && c <= '9') {
            value = c - '0';
            return true;
        }
        if (c >= 'a' && c <= 'f') {
            value = 10 + (c - 'a');
            return true;
        }
        if (c >= 'A' && c <= 'F') {
            value = 10 + (c - 'A');
            return true;
        }

        value = 0;
        return false;
    }

    private static int FindObjectEnd(string text, int start) {
        int searchFrom = start;
        while (searchFrom >= 0 && searchFrom < text.Length) {
            int streamIdx = IndexOfKeywordOutsideLiteralString(text, "stream", searchFrom, text.Length);
            int endObjIdx = IndexOfKeywordOutsideLiteralString(text, "endobj", searchFrom, text.Length);

            if (endObjIdx < 0) {
                return -1;
            }

            if (streamIdx < 0 || endObjIdx < streamIdx) {
                return endObjIdx + 6;
            }

            int afterStream = SkipEOL(text, streamIdx + 6, text.Length);
            int endStreamIdx = IndexOfKeyword(text, "endstream", afterStream, text.Length);
            if (endStreamIdx < 0) {
                return -1;
            }

            searchFrom = endStreamIdx + 9;
        }

        return -1;
    }

    private static int IndexOfKeywordOutsideLiteralString(string text, string keyword, int start, int limit) {
        if (start < 0) start = 0;
        if (limit > text.Length) limit = text.Length;

        int literalDepth = 0;
        bool escaped = false;
        for (int i = start; i < limit; i++) {
            char c = text[i];
            if (literalDepth > 0) {
                if (escaped) {
                    escaped = false;
                    continue;
                }

                if (c == '\\') {
                    escaped = true;
                    continue;
                }

                if (c == '(') {
                    literalDepth++;
                    continue;
                }

                if (c == ')') {
                    literalDepth--;
                }

                continue;
            }

            if (c == '(') {
                literalDepth = 1;
                continue;
            }

            if (i + keyword.Length <= limit &&
                string.CompareOrdinal(text, i, keyword, 0, keyword.Length) == 0 &&
                HasKeywordBoundary(text, i - 1, start, limit) &&
                HasKeywordBoundary(text, i + keyword.Length, start, limit)) {
                return i;
            }
        }

        return -1;
    }

    private static int FindDictEnd(string text, int dictStart, int limit) {
        int depth = 0;
        for (int i = dictStart; i + 1 < limit; i++) {
            char c = text[i]; char n = text[i + 1];
            if (c == '<' && n == '<') { depth++; i++; continue; }
            if (c == '>' && n == '>') { depth--; i++; if (depth == 0) return i + 1; continue; }
        }
        return -1;
    }

    private static int IndexOfKeyword(string text, string keyword, int start, int limit) {
        if (start < 0) start = 0;
        if (limit > text.Length) limit = text.Length;

        int searchFrom = start;
        while (searchFrom < limit) {
            int idx = text.IndexOf(keyword, searchFrom, StringComparison.Ordinal);
            if (idx < 0 || idx >= limit) {
                return -1;
            }

            int end = idx + keyword.Length;
            if (HasKeywordBoundary(text, idx - 1, start, limit) &&
                HasKeywordBoundary(text, end, start, limit)) {
                return idx;
            }

            searchFrom = idx + 1;
        }

        return -1;
    }

    private static bool HasKeywordBoundary(string text, int idx, int start, int limit) {
        if (idx < start || idx >= limit) {
            return true;
        }

        char c = text[idx];
        return char.IsWhiteSpace(c) || c is '<' or '>' or '[' or ']' or '%';
    }

    private static int SkipEOL(string text, int idx, int limit) {
        if (idx < limit) {
            if (text[idx] == '\r') idx++;
            if (idx < limit && text[idx] == '\n') idx++;
        }
        return idx;
    }

    private static string SafeSlice(string s, int start, int length, int maxLen) {
        int len = Math.Min(length, Math.Max(0, maxLen));
        if (start < 0) start = 0;
        if (start + len > s.Length) len = s.Length - start;
        if (len <= 0) return string.Empty;
        return s.Substring(start, len);
    }
}
