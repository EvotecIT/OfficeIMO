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

    internal static IReadOnlyList<string> FingerprintImageData(Dictionary<int, PdfIndirectObject> objects, PdfStream image) {
        var fingerprints = new List<string> { PdfRedactionPlan.ComputeSourceSha256(image.Data) };
        foreach (string key in new[] { "SMask", "Mask" }) {
            if (image.Dictionary.Items.TryGetValue(key, out PdfObject? value) && PdfObjectLookup.Resolve(objects, value) is PdfStream mask)
                fingerprints.Add(PdfRedactionPlan.ComputeSourceSha256(mask.Data));
        }
        return fingerprints;
    }

    private static PdfDictionary? AsDictionary(PdfObject value) =>
        value is PdfDictionary direct ? direct : value is PdfStream stream ? stream.Dictionary : null;

    // Rewriting a repeated image or form paints a scrubbed copy under a new name, and the original stays
    // listed. Pruning keeps anything listed, so its pixels would stay in the file unless the entry goes too.
    // A source file can also arrive with an unpainted form holding searched text. Nothing shows it and
    // nothing would ever need it, so it goes the same way.
    private static bool RemoveUnpaintedRedactedResources(Dictionary<int, PdfIndirectObject> objects, IEnumerable<int> imageObjectNumbers, PdfRedactionArea[] areas, PdfReadLimits limits, HashSet<PdfStream> sourceStreamIdentities) {
        var images = new HashSet<int>(imageObjectNumbers);
        PdfRedactionSearchCriteria[] criteria = GetSearchCriteria(areas);
        if (images.Count == 0 && criteria.Length == 0) return false;
        PdfDictionary? anyPage = objects.Values.Select(static item => item.Value).OfType<PdfDictionary>()
            .FirstOrDefault(static candidate => string.Equals(candidate.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal));
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
                if (owner.Value.Contains(entry.Key) || inheritedNames.Contains(entry.Key)) continue;
                if (!HoldsRedactedImage(entry.Value) && !HoldsSearched(entry.Value, owner.Key)) continue;
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

        bool HoldsSearched(PdfObject value, PdfDictionary ownerResources) {
            if (criteria.Length == 0 || anyPage is null || PdfObjectLookup.Resolve(objects, value) is not PdfStream form ||
                !string.Equals(form.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) return false;
            PdfDictionary formResources = ResolveDictionary(objects, form.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null) ?? ownerResources;
            return StreamHoldsSearchedText(objects, form, formResources, anyPage, criteria, limits) ||
                ResolveDictionary(objects, formResources.Items.TryGetValue("XObject", out PdfObject? nested) ? nested : null) is PdfDictionary xObjects &&
                !ReferenceEquals(formResources, ownerResources) && xObjects.Items.Values.Any(child => HoldsSearched(child, formResources));
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
        PdfRedactionMatch[] reviewedImages = plan.Matches
            .Where(static match => match.Kind == PdfRedactionMatchKind.ImagePlacement && match.RequiresRemoval).ToArray();
        var redactedImages = new HashSet<string>(reviewedImages.SelectMany(static match => match.ImageDataSha256), StringComparer.Ordinal);
        bool reviewedInlineImage = reviewedImages.Any(static match => match.ObjectNumber is null);
        if (criteria.Length == 0 && redactedImages.Count == 0 && !reviewedInlineImage) return Array.Empty<PdfRedactionVerificationIssue>();
        var reviewedPages = new HashSet<int>(plan.Areas.Select(static area => area.PageNumber));
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

        PdfRedactionFormAnalysis forms = AnalyzeFormFields(document, plan.Areas);
        foreach (PdfRedactionAffectedField field in forms.Fields) {
            issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "FormFieldValue", "A form field still holds searched text: " + field.Name + "."));
        }
        if (forms.XfaHoldsSearch) {
            issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "XfaData", "XFA form data still holds searched text."));
        }

        var painted = new HashSet<PdfStream>();
        var listed = new Dictionary<PdfStream, PdfDictionary>();
        var visited = new HashSet<PdfStream>();
        var listedImages = new HashSet<string>(StringComparer.Ordinal);
        var paintedImages = new HashSet<string>(StringComparer.Ordinal);
        PdfDictionary? pageForDecoding = null;
        bool holdersInScope = false;
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            PdfReadPage page = document.Pages[pageIndex];
            cancellationToken.ThrowIfCancellationRequested();
            // Patterns and appearances paint on their own pages, which a page-limited plan left alone.
            holdersInScope = reviewedPages.Contains(pageIndex + 1);
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
            if (holdersInScope) CollectAnnotationAppearances(pageDictionary, resources);
            if (pageForDecoding is null) pageForDecoding = pageDictionary;
        }

        // A form listed on one page can be painted by another, so this waits until every page is walked.
        foreach (KeyValuePair<PdfStream, PdfDictionary> candidate in listed) {
            cancellationToken.ThrowIfCancellationRequested();
            if (criteria.Length == 0 || painted.Contains(candidate.Key) || pageForDecoding is null || !StreamHoldsSearchedText(objects, candidate.Key, candidate.Value, pageForDecoding, criteria, limits)) continue;
            issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "OrphanedFormXObject",
                "Content that the page text doesn't show still holds searched text: a form no page paints, a pattern, a soft mask or an annotation appearance."));
        }
        if (reviewedInlineImage && listed.Keys.Any(form => !painted.Contains(form) && TryDecodeContent(form, objects, limits, out string formContent) && HasInlineImage(formContent, limits))) {
            issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "OrphanedImageData",
                "An inline image is still in the file, held by a form that no page paints, and an inline image was reviewed for redaction."));
        }
        // Identical data still painted elsewhere is a placement the plan left alone, not a leftover.
        foreach (string image in listedImages) {
            if (paintedImages.Contains(image)) continue;
            issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "OrphanedImageData",
                "Image data reviewed for redaction is still in the file, held by a resource that no page paints."));
        }
        return issues;

        void CollectForms(string content, PdfDictionary resources, bool paintedOwner) {
            PdfDictionary xObjects = ResolveDictionary(objects, resources.Items.TryGetValue("XObject", out PdfObject? xObjectsValue) ? xObjectsValue : null) ?? new PdfDictionary();
            var invoked = new HashSet<string>(StringComparer.Ordinal);
            PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations, operation => {
                if (operation.Name == "Do" && operation.Operands.Count > 0 && operation.Operands[operation.Operands.Count - 1] is string name) invoked.Add(name);
            }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
            foreach (KeyValuePair<string, PdfObject> entry in xObjects.Items) {
                if (PdfObjectLookup.Resolve(objects, entry.Value) is not PdfStream form) continue;
                string? subtype = form.Dictionary.Get<PdfName>("Subtype")?.Name;
                bool paints = paintedOwner && invoked.Contains(entry.Key);
                if (string.Equals(subtype, "Image", StringComparison.Ordinal) && redactedImages.Count > 0) {
                    foreach (string fingerprint in FingerprintImageData(objects, form)) {
                        if (redactedImages.Contains(fingerprint)) (paints ? paintedImages : listedImages).Add(fingerprint);
                    }
                }
                if (!string.Equals(subtype, "Form", StringComparison.Ordinal)) continue;
                bool firstVisit = visited.Add(form);
                bool newlyPainted = paints && painted.Add(form);
                PdfDictionary formResources = ResolveDictionary(objects, form.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null) ?? resources;
                listed[form] = formResources;
                // Unreadable content proves nothing is painted, so its resources still count as held.
                if (firstVisit || newlyPainted) CollectForms(TryDecodeContent(form, objects, limits, out string formContent) ? formContent : string.Empty, formResources, paints);
            }
            if (!holdersInScope) return;
            if (ResolveDictionary(objects, resources.Items.TryGetValue("Pattern", out PdfObject? patternsValue) ? patternsValue : null) is PdfDictionary patterns) {
                foreach (PdfObject pattern in patterns.Items.Values) {
                    if (PdfObjectLookup.Resolve(objects, pattern) is PdfStream tiling) CollectHolder(tiling, resources);
                }
            }
            if (ResolveDictionary(objects, resources.Items.TryGetValue("ExtGState", out PdfObject? statesValue) ? statesValue : null) is PdfDictionary states) {
                foreach (PdfObject state in states.Items.Values) {
                    if (ResolveDictionary(objects, state) is PdfDictionary graphicsState &&
                        ResolveDictionary(objects, graphicsState.Items.TryGetValue("SMask", out PdfObject? softMaskValue) ? softMaskValue : null) is PdfDictionary softMask &&
                        PdfObjectLookup.Resolve(objects, softMask.Items.TryGetValue("G", out PdfObject? group) ? group : null) is PdfStream groupForm) CollectHolder(groupForm, resources);
                }
            }
        }

        // These paint outside the reviewed placements, so what they hold isn't covered by the rewrite.
        void CollectHolder(PdfStream holder, PdfDictionary inherited) {
            if (!visited.Add(holder)) return;
            PdfDictionary holderResources = ResolveDictionary(objects, holder.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null) ?? inherited;
            listed[holder] = holderResources;
            CollectForms(TryDecodeContent(holder, objects, limits, out string holderContent) ? holderContent : string.Empty, holderResources, paintedOwner: false);
        }

        void CollectAnnotationAppearances(PdfDictionary pageDictionary, PdfDictionary resources) {
            if (PdfObjectLookup.Resolve(objects, pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotsValue) ? annotsValue : null) is not PdfArray annotations) return;
            foreach (PdfObject item in annotations.Items) {
                if (ResolveDictionary(objects, item) is not PdfDictionary annotation ||
                    string.Equals(annotation.Get<PdfName>("Subtype")?.Name, "Widget", StringComparison.Ordinal)) continue;
                if (criteria.Length > 0 && (HoldsSearchedText(objects, annotation, "Contents", criteria) || HoldsSearchedText(objects, annotation, "RC", criteria))) {
                    issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "AnnotationText", "An annotation on a reviewed page still holds searched text."));
                }
                foreach (PdfStream appearance in GetAppearanceStreams(objects, annotation)) CollectHolder(appearance, resources);
            }
        }

    }

    private static bool StreamHoldsSearchedText(Dictionary<int, PdfIndirectObject> objects, PdfStream form, PdfDictionary formResources,
        PdfDictionary pageDictionary, PdfRedactionSearchCriteria[] criteria, PdfReadLimits limits) {
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
            // Text that can't be read can't be proven free of the search.
            return true;
        }
    }

    /// <summary>Finds the forms listed for a page that it never paints and that hold searched text, which is what the rewrite clears.</summary>
    internal static HashSet<PdfStream> FindUnpaintedFormsHoldingSearch(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? page, IEnumerable<PdfRedactionArea> areas) {
        var found = new HashSet<PdfStream>();
        PdfRedactionSearchCriteria[] criteria = GetSearchCriteria(areas);
        if (criteria.Length == 0 || page is null || GetInheritedDictionary(objects, page, "Resources") is not PdfDictionary pageResources) return found;
        var limits = new PdfReadLimits();
        var painted = new HashSet<PdfStream>();
        var candidates = new List<(PdfStream Form, PdfDictionary Resources)>();
        var content = new StringBuilder();
        if (page.Items.TryGetValue("Contents", out PdfObject? contents)) {
            foreach (PdfReference reference in EnumerateContentStreamReferences(objects, contents)) {
                if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) && indirect.Value is PdfStream stream) {
                    if (!TryDecodeContent(stream, objects, limits, out string segment)) return found;
                    content.Append(segment).Append('\n');
                }
            }
        }
        Visit(content.ToString(), pageResources);
        foreach ((PdfStream form, PdfDictionary resources) in candidates) {
            if (!painted.Contains(form) && StreamHoldsSearchedText(objects, form, resources, page, criteria, limits)) found.Add(form);
        }
        return found;

        void Visit(string ownerContent, PdfDictionary resources) {
            if (ResolveDictionary(objects, resources.Items.TryGetValue("XObject", out PdfObject? value) ? value : null) is not PdfDictionary xObjects) return;
            var invoked = new HashSet<string>(TextContentParser.ExtractFormInvocations(ownerContent, maxOperations: limits.MaxContentOperations,
                maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands).Select(static invocation => invocation.Name), StringComparer.Ordinal);
            foreach (KeyValuePair<string, PdfObject> entry in xObjects.Items) {
                if (PdfObjectLookup.Resolve(objects, entry.Value) is not PdfStream form ||
                    !string.Equals(form.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal)) continue;
                PdfDictionary formResources = ResolveDictionary(objects, form.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null) ?? resources;
                if (!invoked.Contains(entry.Key)) candidates.Add((form, formResources));
                else if (painted.Add(form) && TryDecodeContent(form, objects, limits, out string formContent)) Visit(formContent, formResources);
            }
        }
    }

    /// <summary>True when an annotation holds the searched text in its contents or appearance, wherever it sits on the page.</summary>
    /// <remarks>Form widgets are left out: their values belong to the form field, not to page text.</remarks>
    internal static bool CarriesSearchedText(Dictionary<int, PdfIndirectObject> objects, int pageObjectNumber, PdfAnnotation annotation, IEnumerable<PdfRedactionArea> areas) {
        return CarriesSearchedText(objects, pageObjectNumber, annotation, GetSearchCriteria(areas));
    }

    internal static bool CarriesSearchedText(Dictionary<int, PdfIndirectObject> objects, int pageObjectNumber, PdfAnnotation annotation, PdfRedactionSearchCriteria[] criteria) {
        if (criteria.Length == 0 || string.Equals(annotation.Subtype, "Widget", StringComparison.Ordinal)) return false;
        if (annotation.Contents is string contents && criteria.Any(criterion => criterion.Matches(contents))) return true;
        return annotation.ObjectNumber is int number && objects.TryGetValue(number, out PdfIndirectObject? indirect) && indirect.Value is PdfDictionary dictionary &&
            objects.TryGetValue(pageObjectNumber, out PdfIndirectObject? pageObject) && pageObject.Value is PdfDictionary page &&
            AnnotationHoldsSearchedText(objects, dictionary, page, criteria, new PdfReadLimits());
    }

    private static bool AnnotationHoldsSearchedText(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation, PdfDictionary page, PdfRedactionSearchCriteria[] criteria, PdfReadLimits limits) {
        if (HoldsSearchedText(objects, annotation, "Contents", criteria) || HoldsSearchedText(objects, annotation, "RC", criteria)) return true;
        PdfDictionary pageResources = GetInheritedDictionary(objects, page, "Resources") ?? new PdfDictionary();
        return GetAppearanceStreams(objects, annotation).Any(appearance => StreamHoldsSearchedText(objects, appearance,
            ResolveDictionary(objects, appearance.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null) ?? pageResources, page, criteria, limits));
    }

    internal static IEnumerable<PdfStream> GetAppearanceStreams(Dictionary<int, PdfIndirectObject> objects, PdfDictionary annotation) {
        if (ResolveDictionary(objects, annotation.Items.TryGetValue("AP", out PdfObject? value) ? value : null) is not PdfDictionary appearances) yield break;
        foreach (PdfObject appearance in appearances.Items.Values) {
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, appearance);
            if (resolved is PdfStream stream) yield return stream;
            else if (resolved is PdfDictionary states) {
                foreach (PdfObject state in states.Items.Values) {
                    if (PdfObjectLookup.Resolve(objects, state) is PdfStream stateStream) yield return stateStream;
                }
            }
        }
    }

    private static bool HasInlineImage(string content, PdfReadLimits limits) {
        bool found = false;
        PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations, operation => found |= operation.Name is "BI" or "ID" or "EI",
            maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
        return found;
    }

    private static bool TryDecodeContent(PdfStream stream, Dictionary<int, PdfIndirectObject> objects, PdfReadLimits limits, out string content) {
        content = string.Empty;
        if (stream.DecodingFailed || !StreamDecoder.TryDecode(stream, limits.MaxDecodedStreamBytes, out byte[] decoded, objects)) return false;
        content = PdfEncoding.Latin1GetString(decoded);
        return true;
    }
}
