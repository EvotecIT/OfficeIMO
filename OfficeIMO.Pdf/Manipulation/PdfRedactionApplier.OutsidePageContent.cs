using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private static readonly string[] AlternateTextKeys = { "Alt", "ActualText", "E", "TU" };

    private static PdfRedactionSearchCriteria[] GetSearchCriteria(IEnumerable<PdfRedactionArea> areas) =>
        areas.Select(static area => area.SearchCriteria).OfType<PdfRedactionSearchCriteria>().Distinct().ToArray();

    // Tagged alternate text repeats wording outside the page content. Search marks remove the entries
    // holding what was searched, whatever the cleanup scope, and leave every other description in place.
    private static bool RemoveAlternateTextHoldingMatches(Dictionary<int, PdfIndirectObject> objects, PdfRedactionArea[] areas) {
        PdfRedactionSearchCriteria[] criteria = GetSearchCriteria(areas);
        if (criteria.Length == 0) return false;
        bool changed = false;
        foreach (PdfIndirectObject item in objects.Values) changed = RemoveMatchingAlternateText(objects, item.Value, criteria) || changed;
        return changed;
    }

    private static bool RemoveMatchingAlternateText(Dictionary<int, PdfIndirectObject> objects, PdfObject value, PdfRedactionSearchCriteria[] criteria) {
        PdfDictionary? dictionary = AsDictionary(value);
        if (dictionary is null) return false;
        bool changed = false;
        foreach (string key in AlternateTextKeys) {
            if (HoldsSearchedText(objects, dictionary, key, criteria)) changed = dictionary.Items.Remove(key) || changed;
        }
        foreach (PdfObject child in dictionary.Items.Values.ToArray()) changed = RemoveMatchingAlternateText(objects, child, criteria) || changed;
        return changed;
    }

    private static bool HoldsSearchedAlternateText(Dictionary<int, PdfIndirectObject> objects, PdfObject value, PdfRedactionSearchCriteria[] criteria) {
        PdfDictionary? dictionary = AsDictionary(value);
        if (dictionary is null) return false;
        return AlternateTextKeys.Any(key => HoldsSearchedText(objects, dictionary, key, criteria)) ||
            dictionary.Items.Values.Any(child => HoldsSearchedAlternateText(objects, child, criteria));
    }

    private static bool HoldsSearchedText(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key, PdfRedactionSearchCriteria[] criteria) =>
        dictionary.Items.TryGetValue(key, out PdfObject? entry) && PdfObjectLookup.Resolve(objects, entry) is PdfStringObj text &&
        criteria.Any(criterion => criterion.Matches(text.Value));

    private static PdfDictionary? AsDictionary(PdfObject value) =>
        value is PdfDictionary direct ? direct : value is PdfStream stream ? stream.Dictionary : null;

    /// <summary>Finds searched text the page extractors cannot see: alternate text, and Form XObjects kept in resources but never painted.</summary>
    internal static IReadOnlyList<PdfRedactionVerificationIssue> FindSearchedTextOutsidePageContent(
        byte[] pdf, IEnumerable<PdfRedactionArea> areas, PdfLoadOptions? readOptions, System.Threading.CancellationToken cancellationToken) {
        PdfRedactionSearchCriteria[] criteria = GetSearchCriteria(areas);
        if (criteria.Length == 0) return Array.Empty<PdfRedactionVerificationIssue>();
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        PdfReadLimits limits = readOptions?.Limits ?? new PdfReadLimits();
        var issues = new List<PdfRedactionVerificationIssue>();
        foreach (PdfIndirectObject item in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!HoldsSearchedAlternateText(objects, item.Value, criteria)) continue;
            issues.Add(new PdfRedactionVerificationIssue("SearchedTextOutsidePageContent", "AlternateText",
                "Alternate text in object " + item.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " still holds searched text."));
        }

        var painted = new HashSet<PdfStream>();
        var listed = new Dictionary<PdfStream, PdfDictionary>();
        var visited = new HashSet<PdfStream>();
        PdfDictionary? pageForDecoding = null;
        foreach (PdfReadPage page in document.Pages) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!objects.TryGetValue(page.ObjectNumber, out PdfIndirectObject? pageObject) || pageObject.Value is not PdfDictionary pageDictionary ||
                GetInheritedDictionary(objects, pageDictionary, "Resources") is not PdfDictionary resources) continue;
            var content = new StringBuilder();
            if (pageDictionary.Items.TryGetValue("Contents", out PdfObject? contents)) {
                foreach (PdfReference reference in EnumerateContentReferences(objects, contents)) {
                    if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) && indirect.Value is PdfStream stream &&
                        TryDecodeContent(stream, objects, limits, out string segment)) content.Append(segment).Append('\n');
                }
            }
            CollectForms(content.ToString(), resources, paintedOwner: true);
            if (pageForDecoding is null) pageForDecoding = pageDictionary;
        }

        // A form listed on one page can be painted by another, so this waits until every page is walked.
        foreach (KeyValuePair<PdfStream, PdfDictionary> candidate in listed) {
            cancellationToken.ThrowIfCancellationRequested();
            if (painted.Contains(candidate.Key) || pageForDecoding is null || !HoldsSearchedText(candidate.Key, candidate.Value, pageForDecoding)) continue;
            issues.Add(new PdfRedactionVerificationIssue("SearchedTextOutsidePageContent", "OrphanedFormXObject",
                "A Form XObject that no page paints still holds searched text."));
        }
        return issues;

        void CollectForms(string content, PdfDictionary resources, bool paintedOwner) {
            if (ResolveDictionary(objects, resources.Items.TryGetValue("XObject", out PdfObject? xObjectsValue) ? xObjectsValue : null) is not PdfDictionary xObjects) return;
            var invoked = new HashSet<string>(StringComparer.Ordinal);
            PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations, operation => {
                if (operation.Name == "Do" && operation.Operands.Count > 0 && operation.Operands[operation.Operands.Count - 1] is string name) invoked.Add(name);
            }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
            foreach (KeyValuePair<string, PdfObject> entry in xObjects.Items) {
                if (PdfObjectLookup.Resolve(objects, entry.Value) is not PdfStream form ||
                    !string.Equals(form.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) continue;
                bool paints = paintedOwner && invoked.Contains(entry.Key);
                bool firstVisit = visited.Add(form);
                bool newlyPainted = paints && painted.Add(form);
                PdfDictionary formResources = ResolveDictionary(objects, form.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null) ?? resources;
                listed[form] = formResources;
                if ((firstVisit || newlyPainted) && TryDecodeContent(form, objects, limits, out string formContent))
                    CollectForms(formContent, formResources, paints);
            }
        }

        bool HoldsSearchedText(PdfStream form, PdfDictionary formResources, PdfDictionary pageDictionary) {
            if (!TryDecodeContent(form, objects, limits, out string formContent)) return false;
            try {
                Dictionary<string, Func<byte[], string>> decoders = MergeDecoders(ResourceResolver.GetFontDecoders(pageDictionary, objects),
                    ResourceResolver.GetFontDecodersForResources(formResources, objects));
                Dictionary<string, Func<byte[], double>> widths = MergeWidthProviders(ResourceResolver.GetFontWidthProviders(pageDictionary, objects),
                    ResourceResolver.GetFontWidthProvidersForResources(formResources, objects));
                string[] runs = ParseTextSpans(formContent, decoders, widths, limits).Select(static span => span.Text).ToArray();
                string joined = string.Concat(runs), spaced = string.Join(" ", runs);
                return criteria.Any(criterion => criterion.Matches(joined) || criterion.Matches(spaced));
            } catch (PdfUnsupportedTextMappingException) {
                // Text that can't be read can't be cleared, and nothing paints this form anyway.
                return true;
            }
        }
    }

    private static bool TryDecodeContent(PdfStream stream, Dictionary<int, PdfIndirectObject> objects, PdfReadLimits limits, out string content) {
        content = string.Empty;
        if (stream.DecodingFailed || !StreamDecoder.TryDecode(stream, limits.MaxDecodedStreamBytes, out byte[] decoded, objects)) return false;
        content = PdfEncoding.Latin1GetString(decoded);
        return true;
    }
}
