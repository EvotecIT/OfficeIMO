using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private static readonly string[] AlternateTextKeys = { "Alt", "ActualText", "E", "TU" };

    private static PdfRedactionSearchCriteria[] GetSearchCriteria(IEnumerable<PdfRedactionArea> areas) =>
        areas.Select(static area => area.SearchCriteria).OfType<PdfRedactionSearchCriteria>().Distinct().ToArray();

    // Tagged alternate text repeats wording outside the page content. Search marks remove the entries
    // holding what was searched, whatever the cleanup scope, and leave every other description in place.
    private static bool RemoveAlternateTextHoldingMatches(Dictionary<int, PdfIndirectObject> objects, PdfRedactionArea[] areas, PdfReadLimits limits) {
        PdfRedactionSearchCriteria[] criteria = GetSearchCriteria(areas);
        if (criteria.Length == 0) return false;
        bool changed = false;
        foreach (PdfIndirectObject item in objects.Values) changed = RemoveMatchingAlternateText(objects, item.Value, criteria) || changed;
        HashSet<int> pageContentStreams = GetPageContentStreamNumbers(objects);
        foreach (PdfIndirectObject item in objects.Values.ToArray()) {
            if (!IsContentOwner(pageContentStreams, item, out PdfStream? stream) || stream.DecodingFailed ||
                !StreamDecoder.TryDecode(stream, int.MaxValue, out byte[] decoded, objects)) continue;
            string content = PdfEncoding.Latin1GetString(decoded);
            string rewritten = RewriteMarkedContentHoldingMatches(content, criteria, limits, out _);
            if (string.Equals(content, rewritten, StringComparison.Ordinal)) continue;
            objects[item.ObjectNumber] = new PdfIndirectObject(item.ObjectNumber, item.Generation,
                new PdfStream(CleanStreamDictionary(stream.Dictionary), PdfEncoding.Latin1GetBytes(rewritten)));
            changed = true;
        }
        return changed;
    }

    private static HashSet<int> GetPageContentStreamNumbers(Dictionary<int, PdfIndirectObject> objects) {
        var numbers = new HashSet<int>();
        foreach (PdfIndirectObject candidate in objects.Values) {
            if (candidate.Value is not PdfDictionary page || !string.Equals(page.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal) ||
                !page.Items.TryGetValue("Contents", out PdfObject? contents)) continue;
            foreach (PdfReference reference in EnumerateContentStreamReferences(objects, contents)) numbers.Add(reference.ObjectNumber);
        }
        return numbers;
    }

    private static bool IsContentOwner(HashSet<int> pageContentStreams, PdfIndirectObject item, out PdfStream stream) {
        stream = (item.Value as PdfStream)!;
        return stream is not null && (pageContentStreams.Contains(item.ObjectNumber) ||
            string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal));
    }

    // Marked content can carry text inline, beside the glyphs rather than in them. Only the property list goes;
    // the section stays open so its matching EMC still balances.
    private static string RewriteMarkedContentHoldingMatches(string content, PdfRedactionSearchCriteria[] criteria, PdfReadLimits limits, out bool found) {
        var edits = new List<(int Start, int End)>();
        int previousEnd = 0;
        PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations, operation => {
            int start = PdfContentStreamTextRewriter.FindOperationStart(content, previousEnd, operation.OperatorOffset);
            previousEnd = Math.Min(content.Length, operation.OperatorOffset + operation.Name.Length);
            if (operation.Name != "BDC" || operation.Operands.Count == 0 ||
                operation.Operands[operation.Operands.Count - 1] is not PdfContentDictionary dictionary) return;
            if (AlternateTextKeys.Any(key => dictionary.Items.TryGetValue(key, out object? value) && value is byte[] bytes &&
                criteria.Any(criterion => criterion.Matches(PdfTextString.Decode(bytes))))) edits.Add((start, previousEnd));
        }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
        found = edits.Count > 0;
        if (!found) return content;
        var builder = new StringBuilder(content.Length);
        int cursor = 0;
        foreach ((int start, int end) in edits) {
            builder.Append(content, cursor, start - cursor).Append("/Span BMC");
            cursor = end;
        }
        return builder.Append(content, cursor, content.Length - cursor).ToString();
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

    // Rewriting a repeated image or form paints a scrubbed copy under a new name, and the original stays
    // listed. Pruning keeps anything listed, so its pixels would stay in the file unless the entry goes too.
    private static bool RemoveUnpaintedRedactedImages(Dictionary<int, PdfIndirectObject> objects, IEnumerable<int> imageObjectNumbers, PdfReadLimits limits, HashSet<PdfStream> sourceStreamIdentities) {
        var images = new HashSet<int>(imageObjectNumbers);
        if (images.Count == 0) return false;
        // A resource dictionary can have several owners, so each collects what all of them paint.
        var paintedNames = new Dictionary<PdfDictionary, HashSet<string>>();
        bool hasUnreadableOwner = false;
        var inheritedNames = new HashSet<string>(StringComparer.Ordinal);
        foreach (PdfIndirectObject item in objects.Values.ToArray()) {
            if (item.Value is PdfDictionary page && string.Equals(page.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal)) {
                if (GetInheritedDictionary(objects, page, "Resources") is not PdfDictionary pageResources || !page.Items.TryGetValue("Contents", out PdfObject? contents)) continue;
                var streams = new List<PdfStream>();
                foreach (PdfReference reference in EnumerateContentStreamReferences(objects, contents)) {
                    if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) && indirect.Value is PdfStream stream) streams.Add(stream);
                }
                AddOwner(pageResources, streams);
            } else if (item.Value is PdfStream stream && (string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal) ||
                stream.Dictionary.Get<PdfNumber>("PatternType")?.Value == 1D)) {
                AddOwner(ResolveDictionary(objects, stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null), new[] { stream });
            } else if ((item.Value as PdfDictionary)?.Get<PdfName>("Subtype")?.Name == "Type3" &&
                ResolveDictionary(objects, ((PdfDictionary)item.Value).Items.TryGetValue("CharProcs", out PdfObject? charProcs) ? charProcs : null) is PdfDictionary procs) {
                AddOwner(ResolveDictionary(objects, ((PdfDictionary)item.Value).Items.TryGetValue("Resources", out PdfObject? own) ? own : null),
                    procs.Items.Values.Select(value => PdfObjectLookup.Resolve(objects, value)).OfType<PdfStream>().ToArray());
            }
        }

        // Content that can't be read may paint any entry, so nothing can be called unpainted.
        if (hasUnreadableOwner) return false;
        var holds = new Dictionary<PdfStream, bool>();
        bool changed = false;
        foreach (KeyValuePair<PdfDictionary, HashSet<string>> owner in paintedNames) {
            if (ResolveDictionary(objects, owner.Key.Items.TryGetValue("XObject", out PdfObject? value) ? value : null) is not PdfDictionary xObjects) continue;
            foreach (KeyValuePair<string, PdfObject> entry in xObjects.Items.ToArray()) {
                if (owner.Value.Contains(entry.Key) || inheritedNames.Contains(entry.Key) || !HoldsRedactedImage(entry.Value)) continue;
                changed = xObjects.Items.Remove(entry.Key) || changed;
            }
        }
        return changed;

        void AddOwner(PdfDictionary? resources, IReadOnlyList<PdfStream> streams) {
            HashSet<string> names = resources is null ? inheritedNames
                : paintedNames.TryGetValue(resources, out HashSet<string>? existing) ? existing : paintedNames[resources] = new HashSet<string>(StringComparer.Ordinal);
            foreach (PdfStream stream in streams) {
                if (stream.DecodingFailed || !StreamDecoder.TryDecode(stream, GetMutationDecodeLimit(stream, limits, sourceStreamIdentities), out byte[] decoded, objects)) {
                    hasUnreadableOwner = true;
                    continue;
                }
                foreach (TextContentParser.FormInvocation invocation in TextContentParser.ExtractFormInvocations(PdfEncoding.Latin1GetString(decoded),
                    maxOperations: limits.MaxContentOperations, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands)) names.Add(invocation.Name);
            }
        }

        bool HoldsRedactedImage(PdfObject value) {
            if (value is PdfReference reference && images.Contains(reference.ObjectNumber)) return true;
            if (PdfObjectLookup.Resolve(objects, value) is not PdfStream stream ||
                !string.Equals(stream.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) return false;
            if (holds.TryGetValue(stream, out bool known)) return known;
            holds[stream] = false;
            bool result = ResolveDictionary(objects, stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null) is PdfDictionary resources &&
                ResolveDictionary(objects, resources.Items.TryGetValue("XObject", out PdfObject? nested) ? nested : null) is PdfDictionary xObjects &&
                xObjects.Items.Values.Any(HoldsRedactedImage);
            holds[stream] = result;
            return result;
        }
    }

    /// <summary>Finds redacted content the page readers cannot see: searched alternate text, and text or reviewed image data kept in resources but never painted.</summary>
    internal static IReadOnlyList<PdfRedactionVerificationIssue> FindRedactedContentOutsidePages(
        byte[] pdf, PdfRedactionPlan plan, PdfLoadOptions? readOptions, System.Threading.CancellationToken cancellationToken) {
        PdfRedactionSearchCriteria[] criteria = GetSearchCriteria(plan.Areas);
        var redactedImages = new HashSet<string>(plan.Matches
            .Where(static match => match.RequiresRemoval)
            .Select(static match => match.ImageDataSha256).OfType<string>(), StringComparer.Ordinal);
        if (criteria.Length == 0 && redactedImages.Count == 0) return Array.Empty<PdfRedactionVerificationIssue>();
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        PdfReadLimits limits = readOptions?.Limits ?? new PdfReadLimits();
        var issues = new List<PdfRedactionVerificationIssue>();
        foreach (PdfIndirectObject item in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!HoldsSearchedAlternateText(objects, item.Value, criteria)) continue;
            issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "AlternateText",
                "Alternate text in object " + item.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " still holds searched text."));
        }

        HashSet<int> pageContentStreams = GetPageContentStreamNumbers(objects);
        foreach (PdfIndirectObject item in objects.Values.ToArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (criteria.Length == 0 || !IsContentOwner(pageContentStreams, item, out PdfStream? stream) || !TryDecodeContent(stream, objects, limits, out string markedContent)) continue;
            RewriteMarkedContentHoldingMatches(markedContent, criteria, limits, out bool found);
            if (found) issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "AlternateText",
                "Marked content in object " + item.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " still holds searched text."));
        }

        var painted = new HashSet<PdfStream>();
        var listed = new Dictionary<PdfStream, PdfDictionary>();
        var visited = new HashSet<PdfStream>();
        var listedImages = new HashSet<string>(StringComparer.Ordinal);
        var paintedImages = new HashSet<string>(StringComparer.Ordinal);
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
            if (criteria.Length == 0 || painted.Contains(candidate.Key) || pageForDecoding is null || !HoldsSearchedText(candidate.Key, candidate.Value, pageForDecoding)) continue;
            issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "OrphanedFormXObject",
                "A Form XObject that no page paints still holds searched text."));
        }
        // Identical data still painted elsewhere is a placement the plan left alone, not a leftover.
        foreach (string image in listedImages) {
            if (paintedImages.Contains(image)) continue;
            issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "OrphanedImageData",
                "Image data reviewed for redaction is still in the file, held by a resource that no page paints."));
        }
        return issues;

        void CollectForms(string content, PdfDictionary resources, bool paintedOwner) {
            if (ResolveDictionary(objects, resources.Items.TryGetValue("XObject", out PdfObject? xObjectsValue) ? xObjectsValue : null) is not PdfDictionary xObjects) return;
            var invoked = new HashSet<string>(StringComparer.Ordinal);
            PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations, operation => {
                if (operation.Name == "Do" && operation.Operands.Count > 0 && operation.Operands[operation.Operands.Count - 1] is string name) invoked.Add(name);
            }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
            foreach (KeyValuePair<string, PdfObject> entry in xObjects.Items) {
                if (PdfObjectLookup.Resolve(objects, entry.Value) is not PdfStream form) continue;
                string? subtype = form.Dictionary.Get<PdfName>("Subtype")?.Name;
                bool paints = paintedOwner && invoked.Contains(entry.Key);
                if (string.Equals(subtype, "Image", StringComparison.Ordinal) && redactedImages.Count > 0) {
                    string fingerprint = PdfRedactionPlan.ComputeSourceSha256(form.Data);
                    if (redactedImages.Contains(fingerprint)) (paints ? paintedImages : listedImages).Add(fingerprint);
                }
                if (!string.Equals(subtype, "Form", StringComparison.Ordinal)) continue;
                bool firstVisit = visited.Add(form);
                bool newlyPainted = paints && painted.Add(form);
                PdfDictionary formResources = ResolveDictionary(objects, form.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null) ?? resources;
                listed[form] = formResources;
                // Unreadable content proves nothing is painted, so its resources still count as held.
                if (firstVisit || newlyPainted) CollectForms(TryDecodeContent(form, objects, limits, out string formContent) ? formContent : string.Empty, formResources, paints);
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
