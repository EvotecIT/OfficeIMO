using System.Threading;
using System.Text.RegularExpressions;
using System.Globalization;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private static readonly string[] FontProgramKeys = { "FontFile", "FontFile2", "FontFile3" };
    /// <summary>Rechecks the saved artifact, including holders unknown to the planner. PDF names are syntax, not string data.</summary>
    internal static IReadOnlyList<PdfRedactionVerificationIssue> FindSearchResidue(byte[] pdf, PdfRedactionPlan plan,
        PdfLoadOptions? readOptions, CancellationToken cancellationToken) {
        PdfRedactionSearchCriteria[] criteria = GetSearchCriteria(plan);
        if (criteria.Length == 0) return Array.Empty<PdfRedactionVerificationIssue>();
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        var issues = new List<PdfRedactionVerificationIssue>();
        Inspect(document, pdf, historical: false);
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
            var visited = new HashSet<PdfObject>();
            foreach (PdfIndirectObject item in current.Objects.Values) {
                cancellationToken.ThrowIfCancellationRequested();
                if (excluded.Contains(item.Value)) continue;
                PdfObject inspected = item.Value;
                // Selected marks may deliberately keep another painted occurrence. The glyph/placement verifier
                // owns that content; the generic sweep still checks its dictionaries and every other holder.
                if (!plan.IsSearchDriven && inspected is PdfStream selectedStream &&
                    (contentStreams.Contains(item.ObjectNumber) || selectedStream.Dictionary.Get<PdfName>("Subtype")?.Name == "Form")) inspected = selectedStream.Dictionary;
                if (SearchObjectHolds(current.Objects, inspected, criteria, current.ReadOptions.Limits, visited, excluded,
                    contentStreams.Contains(item.ObjectNumber), followReferences: false, cancellationToken)) {
                    issues.Add(new PdfRedactionVerificationIssue("SearchResidue", historical ? "HistoricalRevision" : "Object:" + item.ObjectNumber.ToString(CultureInfo.InvariantCulture),
                        historical ? "An earlier PDF revision still holds searched text." : "A saved PDF object still holds searched text."));
                }
            }
            if (!historical) return;
            // Raw historical data can contain strings no longer present in its selected object map.
            // Omit excluded page objects, including their streams, before checking those retained bytes.
            string raw = PdfEncoding.Latin1GetString(bytes);
            foreach (PdfIndirectObject item in current.Objects.Values.Where(item => excluded.Contains(item.Value)).OrderByDescending(static item => item.SourceOffset)) {
                int start = item.SourceOffset;
                if (start < 0) continue;
                int end = raw.IndexOf("endobj", start, StringComparison.Ordinal);
                if (end >= start) raw = raw.Remove(start, end + 6 - start);
            }
            bool encodedResidue = PdfRedactionVerification.ContainsSearchedPdfString(PdfEncoding.Latin1GetBytes(raw), criteria, cancellationToken);
            // Font names and dictionary keys must not masquerade as retained values.
            raw = Regex.Replace(raw, @"/[^\s()<>\[\]{}/%]+", "", RegexOptions.CultureInvariant, RegexTimeout);
            if (encodedResidue || SearchTextHolds(raw, criteria)) issues.Add(new PdfRedactionVerificationIssue("SearchResidue", "HistoricalBytes", "Retained bytes of an earlier PDF revision still hold searched text."));
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
            plan.SearchPageNumbers.Count == 0, plan.SearchPageNumbers);
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
        PdfReadLimits limits, HashSet<PdfObject> visited, HashSet<PdfObject>? excluded = null, bool contentStream = false, bool followReferences = true, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!followReferences && value is PdfReference) return false;
        value = PdfObjectLookup.Resolve(objects, value);
        if (value is null || excluded?.Contains(value) == true || !visited.Add(value)) return false;
        switch (value) {
            case PdfStringObj text: return SearchTextHolds(text.Value, criteria);
            case PdfArray array: return array.Items.Any(item => SearchObjectHolds(objects, item, criteria, limits, visited, excluded, cancellationToken: cancellationToken, followReferences: followReferences));
            case PdfDictionary dictionary: return dictionary.Items.Values.Any(item => SearchObjectHolds(objects, item, criteria, limits, visited, excluded, cancellationToken: cancellationToken, followReferences: followReferences));
            case PdfStream stream:
                if (SearchObjectHolds(objects, stream.Dictionary, criteria, limits, visited, excluded, cancellationToken: cancellationToken, followReferences: followReferences)) return true;
                if (!StreamDecoder.TryDecode(stream, limits.MaxDecodedStreamBytes, out byte[] data, objects)) return false;
                string content = PdfEncoding.Latin1GetString(data);
                bool cmap = ContainsOrdinal(content, "begincmap");
                if (contentStream || stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Form" || stream.Dictionary.Items.ContainsKey("PatternType") || cmap) {
                    bool found = false;
                    PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations,
                        operation => {
                            // A font's repertoire is not retained text. Only actual string data outside its code mappings counts.
                            if (cmap && operation.Name is "endbfchar" or "endbfrange" or "endcodespacerange" or "endcidchar" or "endcidrange") return;
                            found |= operation.Operands.Any(OperandHolds);
                        }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
                    return found;
                }
                // Binary image/font payloads are not text-bearing streams. Their string dictionaries were checked above.
                if (stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image") return false;
                bool fontProgram = objects.Values.Any(item => item.Value is PdfDictionary descriptor &&
                    FontProgramKeys.Any(key => descriptor.Items.TryGetValue(key, out PdfObject? font) && ReferenceEquals(PdfObjectLookup.Resolve(objects, font), stream)));
                return !fontProgram && (SearchTextHolds(content, criteria) || PdfRedactionVerification.ContainsSearchedPdfString(data, criteria, cancellationToken));
            default: return false;
        }

        bool OperandHolds(object operand) => operand switch {
            byte[] bytes => SearchTextHolds(PdfEncoding.Latin1GetString(bytes), criteria),
            List<object> array => array.Any(OperandHolds),
            PdfContentDictionary dictionary => dictionary.Items.Values.Any(OperandHolds),
            _ => false
        };
    }

    /// <summary>Page data exclusively reachable from excluded pages may legitimately retain a searched value.</summary>
    private static HashSet<PdfObject> FindExcludedSearchObjects(Dictionary<int, PdfIndirectObject> objects, int[] pages,
        PdfRedactionSearchCriteria[] criteria, CancellationToken cancellationToken) {
        var included = new HashSet<PdfObject>();
        var excluded = new HashSet<PdfObject>();
        if (criteria.Any(static criterion => criterion.CoversAllPages)) return excluded;
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
        Check(catalog);
        for (int index = 0; index < pages.Count; index++) {
            if (!criteria.Any(criterion => criterion.CoversAllPages || criterion.PageNumbers.Contains(index + 1))) continue;
            if (objects.TryGetValue(pages[index], out PdfIndirectObject? page)) Check(page.Value as PdfDictionary);
        }
        PdfReference? info = PdfSyntax.ReadTrailerReference(trailer, "Info", limits);
        bool removeInfo = info is not null && SearchObjectHolds(objects, info, criteria, limits, new HashSet<PdfObject>());
        return (owners, removeInfo);

        void Check(PdfDictionary? owner) {
            if (owner is not null && owner.Items.TryGetValue("Metadata", out PdfObject? metadata) &&
                SearchObjectHolds(objects, metadata, criteria, limits, new HashSet<PdfObject>())) owners.Add(owner);
        }
    }
}
