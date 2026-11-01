using System.Threading;
using System.Text.RegularExpressions;
using System.Globalization;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private static readonly Regex FileStructure = new Regex(@"(?<!\S)(?:start)?xref\b.*?%%EOF|%PDF-\d\.\d", RegexOptions.Singleline | RegexOptions.CultureInvariant, Regex.InfiniteMatchTimeout);
    private static readonly string[] FontProgramKeys = { "FontFile", "FontFile2", "FontFile3" };
    private static readonly string[] FontCMapKeys = { "ToUnicode", "Encoding", "UseCMap" };
    private static readonly HashSet<string> CMapSyntaxWords = new HashSet<string>(StringComparer.Ordinal) {
        "CIDInit", "CMapName", "ProcSet", "begin", "begincmap", "begincodespacerange", "beginbfchar", "beginbfrange", "begincidchar", "begincidrange",
        "beginnotdefchar", "beginnotdefrange", "currentdict", "def", "defineresource", "dict", "end", "endcmap", "endcodespacerange", "endbfchar",
        "endbfrange", "endcidchar", "endcidrange", "endnotdefchar", "endnotdefrange", "findresource", "pop", "usecmap"
    };
    /// <summary>Rechecks the saved artifact, including holders unknown to the planner. PDF names are syntax, not string data.</summary>
    internal static IReadOnlyList<PdfRedactionVerificationIssue> FindSearchResidue(byte[] pdf, PdfRedactionPlan plan,
        PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        PdfRedactionSearchCriteria[] criteria = GetSearchCriteria(plan);
        if (criteria.Length == 0) return Array.Empty<PdfRedactionVerificationIssue>();
        readOptions = PdfLoadOptions.WithArtifactText(readOptions);
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        var issues = new List<PdfRedactionVerificationIssue>();
        Inspect(document, pdf, historical: false);
        // Do not attest retained compressed containers using a live-object projection alone.
        // Normalized output has no such containers; unproven historical inputs remain findings.
        if (!PdfSyntax.CanNormalizeRedactionObjectGraph(pdf, document, cancellationToken))
            issues.Add(new PdfRedactionVerificationIssue("SearchResidue", "CompressedSource", "Retained compressed source bytes have no proven residue sweep."));
        foreach (int end in PdfSyntax.GetHistoricalRevisionEnds(pdf, document, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            var revision = new byte[end];
            Buffer.BlockCopy(pdf, 0, revision, 0, end);
            Inspect(PdfReadDocument.Open(revision, readOptions, cancellationToken), revision, historical: true);
        }
        return issues;

        void Inspect(PdfReadDocument current, byte[] bytes, bool historical) {
            HashSet<PdfObject> excluded = FindExcludedSearchObjects(current.Objects, current.Pages.Select(static page => page.ObjectNumber).ToArray(), criteria, cancellationToken);
            HashSet<int> contentStreams = GetPageContentStreamNumbers(current.Objects);
            PdfRedactionFontStreams fontStreams = FindFontStreams(current.Objects, cancellationToken);
            if (plan.IsSearchDriven) {
                for (int index = 0; index < current.Pages.Count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    PdfRedactionSearchCriteria[] pageCriteria = criteria.Where(criterion => criterion.CoversAllPages || criterion.PageNumbers.Contains(index + 1)).ToArray();
                    if (pageCriteria.Length == 0) continue;
                    PdfReadPage page = current.Pages[index];
                    if (pageCriteria.Any(criterion => criterion.Matches(page.ExtractText(cancellationToken))) ||
                        page.GetTextSpansIncludingHiddenOptionalContent().Any(span => pageCriteria.Any(criterion => criterion.Matches(span.Text))))
                        issues.Add(new PdfRedactionVerificationIssue("SearchResidue", "Page:" + (index + 1).ToString(CultureInfo.InvariantCulture),
                            "A saved page still paints searched text."));
                }
            }
            var visited = new HashSet<PdfObject>();
            var containerBudget = new PdfDecodedStreamBudget(current.ReadOptions.Limits);
            foreach (PdfIndirectObject item in current.Objects.Values) {
                cancellationToken.ThrowIfCancellationRequested();
                if (excluded.Contains(item.Value)) continue;
                PdfObject inspected = item.Value;
                if (inspected is PdfStream container && container.Dictionary.Get<PdfName>("Type")?.Name == "ObjStm") {
                    foreach (var member in PdfSyntax.ReadRedactionObjectStreamMembers(container, current.Objects, current.ReadOptions.Limits, containerBudget, cancellationToken)) {
                        if (current.Objects.TryGetValue(member.Number, out PdfIndirectObject? live) &&
                            live.SourceOffset == item.SourceOffset && excluded.Contains(live.Value)) continue;
                        if (SearchObjectHolds(current.Objects, member.Value, criteria, current.ReadOptions.Limits, fontStreams,
                            new HashSet<PdfObject>(), excluded, followReferences: false, cancellationToken: cancellationToken))
                            issues.Add(new PdfRedactionVerificationIssue("SearchResidue", "CompressedMember", "A retained compressed PDF member still holds searched text."));
                    }
                    inspected = container.Dictionary;
                }
                if (inspected is PdfStream crossReferences && crossReferences.Dictionary.Get<PdfName>("Type")?.Name == "XRef") inspected = crossReferences.Dictionary;
                // Selected marks may deliberately keep another painted occurrence. The glyph/placement verifier
                // owns that content; the generic sweep still checks its dictionaries and every other holder.
                if (!plan.IsSearchDriven && inspected is PdfStream selectedStream &&
                    (contentStreams.Contains(item.ObjectNumber) || selectedStream.Dictionary.Get<PdfName>("Subtype")?.Name == "Form")) inspected = selectedStream.Dictionary;
                if (SearchObjectHolds(current.Objects, inspected, criteria, current.ReadOptions.Limits, fontStreams, visited, excluded,
                    contentStreams.Contains(item.ObjectNumber), followReferences: false, cancellationToken)) {
                    issues.Add(new PdfRedactionVerificationIssue("SearchResidue", historical ? "HistoricalRevision" : "Object:" + item.ObjectNumber.ToString(CultureInfo.InvariantCulture),
                        historical ? "An earlier PDF revision still holds searched text." : "A saved PDF object still holds searched text."));
                }
            }
            string raw = historical ? RemoveExcludedObjects(bytes, current, excluded) : ReadBytesOutsideObjects(bytes, current, cancellationToken);
            bool encodedResidue = PdfRedactionVerification.ContainsSearchedPdfString(PdfEncoding.Latin1GetBytes(raw), criteria, cancellationToken);
            // Font names and dictionary keys must not masquerade as retained values.
            raw = Regex.Replace(raw, @"/[^\s()<>\[\]{}/%]+", "", RegexOptions.CultureInvariant, RegexTimeout);
            if (encodedResidue || SearchTextHolds(raw, criteria)) issues.Add(historical
                ? new PdfRedactionVerificationIssue("SearchResidue", "HistoricalBytes", "Retained bytes of an earlier PDF revision still hold searched text.")
                : new PdfRedactionVerificationIssue("SearchResidue", "UnparsedBytes", "Bytes outside the PDF's objects still hold searched text."));
        }
    }

    // Raw historical data can contain strings no longer present in its selected object map.
    // Omit excluded page objects, including their streams, before checking those retained bytes.
    private static string RemoveExcludedObjects(byte[] bytes, PdfReadDocument document, HashSet<PdfObject> excluded) {
        string raw = PdfEncoding.Latin1GetString(bytes);
        foreach (PdfIndirectObject item in document.Objects.Values.Where(item => excluded.Contains(item.Value)).OrderByDescending(static item => item.SourceOffset)) {
            int start = item.SourceOffset;
            if (start < 0) continue;
            int end = raw.IndexOf("endobj", start, StringComparison.Ordinal);
            if (end >= start) raw = raw.Remove(start, end + 6 - start);
        }
        return raw;
    }

    // Comments and bytes between or after objects never reach the object map, so no other check sees them.
    // Object bodies were inspected decoded, and their numbers would pass for retained digits, as would the file structure.
    private static string ReadBytesOutsideObjects(byte[] bytes, PdfReadDocument document, CancellationToken cancellationToken) {
        string raw = PdfEncoding.Latin1GetString(bytes);
        var outside = new StringBuilder();
        var streams = new Dictionary<int, bool>();
        foreach (PdfIndirectObject item in document.Objects.Values) if (item.SourceOffset >= 0) streams[item.SourceOffset] = item.Value is PdfStream;
        int cursor = 0;
        foreach (KeyValuePair<int, bool> item in streams.OrderBy(static item => item.Key)) {
            cancellationToken.ThrowIfCancellationRequested();
            int start = item.Key;
            if (start < cursor) continue;
            int bodyEnd = item.Value ? raw.IndexOf("stream", start, StringComparison.Ordinal) : -1;
            int end = raw.IndexOf("endobj", bodyEnd >= 0 ? Math.Max(start, raw.IndexOf("endstream", bodyEnd, StringComparison.Ordinal)) : start, StringComparison.Ordinal);
            if (end < 0) break;
            outside.Append(raw, cursor, start - cursor);
            AppendComments(raw, start, bodyEnd >= 0 ? bodyEnd : end, outside);
            cursor = end + 6;
        }
        outside.Append(raw, cursor, raw.Length - cursor);
        return FileStructure.Replace(outside.ToString(), match => {
            var comments = new StringBuilder();
            AppendComments(match.Value, 0, Math.Max(0, match.Length - 5), comments);
            return comments.ToString();
        });
    }

    private static void AppendComments(string text, int start, int end, StringBuilder target) {
        int index = start;
        while (index < end) {
            if (!TrySkipPdfStringOrComment(text, index, out int next)) { index++; continue; }
            if (text[index] == '%') target.Append(text, index, next - index).Append('\n');
            index = Math.Max(next, index + 1);
        }
    }

    /// <summary>Uses the private criteria even when a no-match plan has no areas carrying them.</summary>
    internal static PdfRedactionSearchCriteria[] GetSearchCriteria(PdfRedactionPlan plan) {
        PdfRedactionSearchCriteria[] criteria = GetSearchCriteria(plan.Areas);
        if (criteria.Length > 0) return criteria;
        PdfRedactionSearchCriteria? requested = PdfRedactionSearchCriteria.Create(
            plan.RequestedSearchCriteria.Where(static value => value.StartsWith("literal:", StringComparison.Ordinal)).Select(static value => value.Substring(8)),
            plan.RequestedSearchCriteria.Where(static value => value.StartsWith("regex:", StringComparison.Ordinal))
                .Select(value => new Regex(value.Substring(6), plan.SearchRegexOptions, plan.SearchRegexTimeout)),
            plan.SearchMatchCase, plan.SearchPageNumbers.Count == 0, plan.SearchPageNumbers);
        return requested is null ? Array.Empty<PdfRedactionSearchCriteria>() : new[] { requested };
    }

    private static bool SearchTextHolds(string text, PdfRedactionSearchCriteria[] criteria) {
        byte[] bytes = PdfEncoding.Latin1GetBytes(text);
        string decoded = PdfTextString.Decode(bytes);
        string utf8 = Encoding.UTF8.GetString(bytes);
        bool hasNull = bytes.Any(static value => value == 0);
        string utf16 = hasNull ? Encoding.Unicode.GetString(bytes) : string.Empty;
        string utf16BigEndian = hasNull ? Encoding.BigEndianUnicode.GetString(bytes) : string.Empty;
        return criteria.Any(criterion => criterion.Matches(text) || criterion.Matches(decoded) || criterion.Matches(utf8) ||
            criterion.Matches(utf16) || criterion.Matches(utf16BigEndian) ||
            criterion.Matches(System.Net.WebUtility.HtmlDecode(XmlTag.Replace(utf8, string.Empty))));
    }

    /// <summary>Checks data recursively, resolving references only when explicitly reached, with stream decoding bounded by the reader limits.</summary>
    private static bool SearchObjectHolds(Dictionary<int, PdfIndirectObject> objects, PdfObject? value, PdfRedactionSearchCriteria[] criteria,
        PdfReadLimits limits, PdfRedactionFontStreams fontStreams, HashSet<PdfObject> visited, HashSet<PdfObject>? excluded = null, bool contentStream = false, bool followReferences = true, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!followReferences && value is PdfReference) return false;
        value = PdfObjectLookup.Resolve(objects, value);
        if (value is null || excluded?.Contains(value) == true || !visited.Add(value)) return false;
        switch (value) {
            case PdfStringObj text: return SearchTextHolds(text.Value, criteria);
            case PdfArray array: return array.Items.Any(item => SearchObjectHolds(objects, item, criteria, limits, fontStreams, visited, excluded, cancellationToken: cancellationToken, followReferences: followReferences));
            case PdfDictionary dictionary: return dictionary.Items.Values.Any(item => SearchObjectHolds(objects, item, criteria, limits, fontStreams, visited, excluded, cancellationToken: cancellationToken, followReferences: followReferences));
            case PdfStream stream:
                if (SearchObjectHolds(objects, stream.Dictionary, criteria, limits, fontStreams, visited, excluded, cancellationToken: cancellationToken, followReferences: followReferences)) return true;
                if (!StreamDecoder.TryDecode(stream, limits.MaxDecodedStreamBytes, out byte[] data, objects)) return false;
                string content = PdfEncoding.Latin1GetString(data);
                // A font's mapping is its repertoire, not retained text. The word alone doesn't make a stream one.
                bool cmap = fontStreams.CMaps.Contains(stream) && ContainsOrdinal(content, "begincmap");
                var cmapWords = new StringBuilder();
                if (contentStream || stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Form" || stream.Dictionary.Items.ContainsKey("PatternType") || cmap) {
                    bool found = false;
                    try {
                        PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations,
                            operation => {
                                if (cmap && !CMapSyntaxWords.Contains(operation.Name)) cmapWords.Append(operation.Name).Append(' ');
                                if (cmap && operation.Name is "endbfchar" or "endbfrange" or "endcidchar" or "endcidrange") return;
                                found |= operation.Operands.Any(OperandHolds);
                            }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
                    } catch (FormatException) {
                        // A sniffed stream can contain arbitrary bytes rather than PDF operands.
                        // Inspect the entire payload, including valid strings after the bad token.
                        return found || SearchTextHolds(content, criteria) || PdfRedactionVerification.ContainsSearchedPdfString(data, criteria, cancellationToken);
                    }
                    return found || SearchTextHolds(cmapWords.ToString(), criteria);
                }
                // Binary image/font payloads are not text-bearing streams. Their string dictionaries were checked above.
                if (stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image") {
                    data = ReadBytesPastRaster(stream, data, objects);
                    content = PdfEncoding.Latin1GetString(data);
                }
                return !fontStreams.Programs.Contains(stream) && (SearchTextHolds(content, criteria) || PdfRedactionVerification.ContainsSearchedPdfString(data, criteria, cancellationToken));
            default: return false;
        }

        bool OperandHolds(object operand) => operand switch {
            byte[] bytes => SearchTextHolds(PdfEncoding.Latin1GetString(bytes), criteria),
            List<object> array => array.Any(OperandHolds),
            PdfContentDictionary dictionary => dictionary.Items.Values.Any(OperandHolds),
            _ => false
        };
    }

    // Pixels can hold any bytes, so they are never searched. A raster's length is declared though, and whatever
    // lies past its last row is never painted, so that part is an ordinary payload. Codec data has no such length.
    private static byte[] ReadBytesPastRaster(PdfStream image, byte[] data, Dictionary<int, PdfIndirectObject> objects) {
        if (!StreamDecoder.TryGetDeclaredFilterNames(image.Dictionary, objects, out IReadOnlyList<string> filters) ||
            filters.Any(static filter => filter is "DCTDecode" or "DCT" or "JPXDecode")) return Array.Empty<byte>();
        double width = PdfObjectLookup.Resolve(objects, image.Dictionary.Items.TryGetValue("Width", out PdfObject? widthValue) ? widthValue : null) is PdfNumber w ? w.Value : 0D;
        double height = PdfObjectLookup.Resolve(objects, image.Dictionary.Items.TryGetValue("Height", out PdfObject? heightValue) ? heightValue : null) is PdfNumber h ? h.Value : 0D;
        if (width < 1D || height < 1D) return data;
        double bitsPerPixel = 1D;
        if (image.Dictionary.Get<PdfBoolean>("ImageMask")?.Value != true) {
            if (!PdfImageColorSpaceNormalization.TryResolve(image.Dictionary.Items.TryGetValue("ColorSpace", out PdfObject? colorSpace) ? colorSpace : null, string.Empty, objects, out PdfImageColorSpaceNormalization normalization) ||
                PdfObjectLookup.Resolve(objects, image.Dictionary.Items.TryGetValue("BitsPerComponent", out PdfObject? bits) ? bits : null) is not PdfNumber bitsPerComponent) return Array.Empty<byte>();
            bitsPerPixel = normalization.SourceColorCount * bitsPerComponent.Value;
        }
        double rasterLength = Math.Ceiling(width * bitsPerPixel / 8D) * height;
        if (rasterLength >= data.Length) return Array.Empty<byte>();
        var excess = new byte[data.Length - (int)rasterLength];
        Buffer.BlockCopy(data, (int)rasterLength, excess, 0, excess.Length);
        return excess;
    }

    // Font dictionaries also sit directly inside resources, so this walks containers as well as indirect objects.
    private static PdfRedactionFontStreams FindFontStreams(Dictionary<int, PdfIndirectObject> objects, CancellationToken cancellationToken) {
        var found = new PdfRedactionFontStreams();
        var visited = new HashSet<PdfObject>();
        var pending = new Stack<PdfObject>(objects.Values.Select(static item => item.Value));
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject value = pending.Pop();
            if (!visited.Add(value)) continue;
            if (value is PdfArray array) foreach (PdfObject item in array.Items) if (item is PdfDictionary or PdfArray) pending.Push(item);
            if (AsDictionary(value) is not PdfDictionary dictionary) continue;
            foreach (KeyValuePair<string, PdfObject> entry in dictionary.Items) {
                if (entry.Value is PdfDictionary or PdfArray) pending.Push(entry.Value);
                HashSet<PdfStream>? role = Array.IndexOf(FontProgramKeys, entry.Key) >= 0 ? found.Programs : Array.IndexOf(FontCMapKeys, entry.Key) >= 0 ? found.CMaps : null;
                if (role is not null && PdfObjectLookup.Resolve(objects, entry.Value) is PdfStream stream) role.Add(stream);
            }
        }
        return found;
    }

    /// <summary>Page data exclusively reachable from excluded pages may legitimately retain a searched value.</summary>
    private static HashSet<PdfObject> FindExcludedSearchObjects(Dictionary<int, PdfIndirectObject> objects, int[] pages,
        PdfRedactionSearchCriteria[] criteria, CancellationToken cancellationToken) {
        var included = new HashSet<PdfObject>();
        var excluded = new HashSet<PdfObject>();
        if (criteria.Length == 0 || criteria.Any(static criterion => criterion.CoversAllPages)) return excluded;
        var scope = new HashSet<int>(criteria.SelectMany(static criterion => criterion.PageNumbers));
        PdfObject? rootPage = null;
        for (int index = 0; index < pages.Length; index++) {
            if (!objects.TryGetValue(pages[index], out PdfIndirectObject? page)) continue;
            rootPage = page.Value;
            Visit(page.Value, scope.Contains(index + 1) ? included : excluded, new HashSet<(PdfObject, bool)>(), false);
        }
        rootPage = null;
        // Only fields in the actual form tree qualify for the widgetless-field exclusion.
        PdfDictionary? catalog = PdfSyntax.FindCatalog(objects, null, cancellationToken: cancellationToken);
        PdfDictionary? form = catalog is not null && catalog.Items.TryGetValue("AcroForm", out PdfObject? formValue) ? ResolveDictionary(objects, formValue) : null;
        var fieldNodes = new HashSet<PdfObject>();
        if (form is not null && form.Items.TryGetValue("Fields", out PdfObject? fields)) ExcludeUnplacedFields(fields);
        // A structure element is reached from the tree, not from its page, so its links say which pages it describes.
        foreach (KeyValuePair<PdfDictionary, HashSet<int>> element in ReadStructureLinks(objects, cancellationToken).Owners) {
            if (element.Value.Count > 0 && element.Value.All(owner => objects.TryGetValue(owner, out PdfIndirectObject? linked) && excluded.Contains(linked.Value))) excluded.Add(element.Key);
        }
        excluded.ExceptWith(included);
        return excluded;

        void ExcludeUnplacedFields(PdfObject? value) {
            value = PdfObjectLookup.Resolve(objects, value);
            if (value is null || !fieldNodes.Add(value)) return;
            if (value is PdfArray array) foreach (PdfObject child in array.Items) ExcludeUnplacedFields(child);
            if (value is not PdfDictionary field) return;
            if (field.Items.TryGetValue("Kids", out PdfObject? kids)) ExcludeUnplacedFields(kids);
            if (!included.Contains(field) && !excluded.Contains(field)) Visit(field, excluded, new HashSet<(PdfObject, bool)>(), true);
        }

        void Visit(PdfObject? value, HashSet<PdfObject> target, HashSet<(PdfObject, bool)> visited, bool ancestor) {
            cancellationToken.ThrowIfCancellationRequested();
            value = PdfObjectLookup.Resolve(objects, value);
            if (value is null || !visited.Add((value, ancestor))) return;
            if (value is PdfDictionary pageTarget && pageTarget.Get<PdfName>("Type")?.Name == "Page" && !ReferenceEquals(value, rootPage)) return;
            target.Add(value);
            if (value is PdfArray array) foreach (PdfObject child in array.Items) Visit(child, target, visited, ancestor);
            PdfDictionary? dictionary = value as PdfDictionary ?? (value as PdfStream)?.Dictionary;
            if (dictionary is null) return;
            target.Add(dictionary);
            foreach (KeyValuePair<string, PdfObject> entry in dictionary.Items) {
                if (entry.Key == "P" || ancestor && entry.Key == "Kids") continue;
                if (entry.Key == "Parent") {
                    if (ResolveDictionary(objects, entry.Value) is PdfDictionary parent && parent.Get<PdfName>("Type")?.Name != "Pages") Visit(parent, target, visited, true);
                    continue;
                }
                Visit(entry.Value, target, visited, ancestor);
            }
        }
    }

    /// <summary>Finds catalog/page XMP and trailer Info that must be removed independently of the optional cleanup policy.</summary>
    internal static (List<PdfDictionary> Owners, bool Info) SearchMetadataHolders(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? catalog,
        string? trailer, IReadOnlyList<int> pages, PdfRedactionSearchCriteria[] criteria, PdfReadLimits limits) {
        var owners = new List<PdfDictionary>();
        if (criteria.Length == 0) return (owners, false);
        PdfRedactionFontStreams fontStreams = FindFontStreams(objects, default);
        Check(catalog);
        for (int index = 0; index < pages.Count; index++) {
            if (!criteria.Any(criterion => criterion.CoversAllPages || criterion.PageNumbers.Contains(index + 1))) continue;
            if (objects.TryGetValue(pages[index], out PdfIndirectObject? page)) Check(page.Value as PdfDictionary);
        }
        PdfReference? info = PdfSyntax.ReadTrailerReference(trailer, "Info", limits);
        bool removeInfo = info is not null && SearchObjectHolds(objects, info, criteria, limits, fontStreams, new HashSet<PdfObject>());
        return (owners, removeInfo);

        void Check(PdfDictionary? owner) {
            if (owner is not null && owner.Items.TryGetValue("Metadata", out PdfObject? metadata) &&
                SearchObjectHolds(objects, metadata, criteria, limits, fontStreams, new HashSet<PdfObject>())) owners.Add(owner);
        }
    }
}

/// <summary>Streams holding a font's own data rather than document text.</summary>
internal sealed class PdfRedactionFontStreams {
    internal HashSet<PdfStream> Programs { get; } = new();
    internal HashSet<PdfStream> CMaps { get; } = new();
}
