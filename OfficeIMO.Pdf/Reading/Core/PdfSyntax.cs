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
    private static readonly Regex TrailerRootRegex = new Regex(@"/Root\s+(\d+)\s+\d+\s+R", RegexOptions.Compiled, RegexTimeout);

    internal static (Dictionary<int, PdfIndirectObject> Map, string TrailerRaw) ParseObjects(byte[] pdf) {
        string text = PdfEncoding.Latin1GetString(pdf);
        var map = new Dictionary<int, PdfIndirectObject>();
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
                    PdfDictionary dict;
                    try { dict = ParseDictionary(dictText); }
                    catch (Exception ex) when (ex is not OutOfMemoryException) { dict = new PdfDictionary(); }

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
                                continue;
                            }
                        }
                    }
                    // No stream; store dictionary-only object
                    map[id] = new PdfIndirectObject(id, gen, dict);
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
                }
            }
        }
        int trailerIdx = text.LastIndexOf("trailer", StringComparison.OrdinalIgnoreCase);
        string trailerRaw = trailerIdx >= 0 ? text.Substring(trailerIdx) : string.Empty;
        ThrowIfEncrypted(trailerRaw);

        ResolveIndirectStreamLengths(map, pdf, streamLocations);
        // Expand object streams (/Type /ObjStm) to populate embedded objects (pages and resources often live there)
        ExpandObjectStreams(map, pdf);
        ThrowIfEncryptedXrefStream(map);
        return (map, trailerRaw);
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
        if (TryGetTrailerRootObjectNumber(trailerRaw, out int rootObjectNumber) &&
            map.TryGetValue(rootObjectNumber, out var rootObject) &&
            rootObject.Value is PdfDictionary rootDictionary &&
            rootDictionary.Get<PdfName>("Type")?.Name == "Catalog") {
            return rootDictionary;
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

    private static bool TryGetTrailerRootObjectNumber(string? trailerRaw, out int objectNumber) {
        objectNumber = 0;
        if (string.IsNullOrWhiteSpace(trailerRaw)) {
            return false;
        }

        Match match = TrailerRootRegex.Match(trailerRaw);
        return match.Success &&
            int.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out objectNumber);
    }

    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> map, PdfObject? value) {
        if (value is PdfReference reference &&
            map.TryGetValue(reference.ObjectNumber, out var indirect)) {
            return indirect.Value;
        }

        return value;
    }

    private static bool IsDestinationForKnownPage(Dictionary<int, PdfIndirectObject> map, PdfArray destination) {
        return destination.Items.Count > 0 &&
            destination.Items[0] is PdfReference pageReference &&
            map.TryGetValue(pageReference.ObjectNumber, out var pageObject) &&
            pageObject.Value is PdfDictionary pageDictionary &&
            pageDictionary.Get<PdfName>("Type")?.Name == "Page";
    }

    private static bool IsDestinationForKnownPage(Dictionary<int, PdfIndirectObject> map, PdfObject destination) {
        return IsDestinationForKnownPage(map, destination, new HashSet<int>());
    }

    private static bool IsDestinationForKnownPage(Dictionary<int, PdfIndirectObject> map, PdfObject destination, HashSet<int> visitedReferences) {
        if (destination is PdfReference reference) {
            if (!visitedReferences.Add(reference.ObjectNumber) ||
                !map.TryGetValue(reference.ObjectNumber, out var indirect)) {
                return false;
            }

            destination = indirect.Value;
        }

        if (destination is PdfArray array) {
            return IsDestinationForKnownPage(map, array);
        }

        if (destination is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("D", out var explicitDestination)) {
            return IsDestinationForKnownPage(map, explicitDestination, visitedReferences);
        }

        return false;
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
            !map.TryGetValue(reference.ObjectNumber, out var indirect) ||
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

                if (!map.TryGetValue(reference.ObjectNumber, out var indirect)) {
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

                if (!map.TryGetValue(reference.ObjectNumber, out var indirect)) {
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
                !map.TryGetValue(reference.ObjectNumber, out var indirect)) {
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

    private static void ExpandObjectStreams(Dictionary<int, PdfIndirectObject> map, byte[] pdf) {
        // Snapshot keys to avoid modifying during enumeration
        var keys = new List<int>(map.Keys);
        foreach (var id in keys) {
            if (!map.TryGetValue(id, out var ind)) continue;
            if (ind.Value is not PdfStream s) continue;
            var type = s.Dictionary.Get<PdfName>("Type")?.Name;
            if (!string.Equals(type, "ObjStm", StringComparison.Ordinal)) continue;

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
                int start = first + off;
                int end = (i + 1 < n) ? first + pairs[i + 1].Off : data.Length;
                if (start < 0 || end > data.Length || end <= start) continue;
                int len = end - start;
                var sliceBytes = new byte[len];
                Buffer.BlockCopy(data, start, sliceBytes, 0, len);
                var slice = PdfEncoding.Latin1GetString(sliceBytes);
                var parsed = ParseTopLevelObject(slice);
                if (parsed is not null) { map[objNum] = new PdfIndirectObject(objNum, 0, parsed); }
            }
        }
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
                    try { return ParseDictionary(dictText); } catch { return new PdfDictionary(); }
                }
            }
            return new PdfDictionary();
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
            return new PdfStringObj(Unescape(inner));
        }
        if (s.Length > 0 && s[0] == '<' && (s.Length == 1 || s[1] != '<')) {
            int end = s.IndexOf('>');
            string inner = end > 1 ? s.Substring(1, end - 1) : s.Substring(1);
            return new PdfStringObj(DecodeHexString(inner));
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
        if (tok.Length > 0 && tok[0] == '(') return (new PdfStringObj(Unescape(tok.Substring(1, tok.Length - 2))), 0);
        if (tok.Length > 1 && tok[0] == '<' && tok[tok.Length - 1] == '>' && (tok.Length == 2 || tok[1] != '<')) {
            return (new PdfStringObj(DecodeHexString(tok.Substring(1, tok.Length - 2))), 0);
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

    private static string Unescape(string s) => PdfTextExtractor.UnescapePdfLiteral(s);

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
        var bytes = DecodeHexBytes(raw);
        if (bytes.Length >= 2) {
            if (bytes[0] == 0xFE && bytes[1] == 0xFF) {
                return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
            }

            if (bytes[0] == 0xFF && bytes[1] == 0xFE) {
                return Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2);
            }
        }

        return PdfWinAnsiEncoding.Decode(bytes);
    }

    private static byte[] DecodeHexBytes(string raw) {
        var hex = new StringBuilder(raw.Length);
        for (int i = 0; i < raw.Length; i++) {
            char ch = raw[i];
            if (!char.IsWhiteSpace(ch)) hex.Append(ch);
        }

        if ((hex.Length & 1) == 1) hex.Append('0');

        var bytes = new byte[hex.Length / 2];
        for (int i = 0; i < bytes.Length; i++) {
            int hi = HexNibble(hex[i * 2]);
            int lo = HexNibble(hex[i * 2 + 1]);
            bytes[i] = (byte)((hi << 4) | lo);
        }

        return bytes;
    }

    private static int HexNibble(char c) {
        if (c >= '0' && c <= '9') return c - '0';
        if (c >= 'a' && c <= 'f') return 10 + (c - 'a');
        if (c >= 'A' && c <= 'F') return 10 + (c - 'A');
        throw new FormatException($"Invalid hex character '{c}'.");
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
