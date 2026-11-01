using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private const string PlainMarkedContent = "/Span BMC";

    private delegate IReadOnlyList<string> PropertyAlternateText(string propertyName);

    private sealed class AlternateTextScope {
        internal AlternateTextScope(int start, int operationEnd, int end, IReadOnlyList<string> values) {
            Start = start;
            OperationEnd = operationEnd;
            End = end;
            Values = values;
        }

        internal int Start { get; }
        internal int OperationEnd { get; }
        internal int End { get; set; }
        internal IReadOnlyList<string> Values { get; }
        internal bool Paints { get; set; }
    }

    // A section can open outside the text object it describes, so this reads the whole stream, not one BT..ET.
    private static List<AlternateTextScope> FindAlternateTextScopes(string content, PropertyAlternateText propertyText, PdfReadLimits limits) {
        var scopes = new List<AlternateTextScope>();
        if (content.IndexOf("BDC", StringComparison.Ordinal) < 0) return scopes;
        var open = new Stack<AlternateTextScope?>();
        int previousEnd = 0;
        PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations, operation => {
            int start = PdfContentStreamTextRewriter.FindOperationStart(content, previousEnd, operation.OperatorOffset);
            previousEnd = Math.Min(content.Length, operation.OperatorOffset + operation.Name.Length);
            if (operation.Name == "BMC") open.Push(null);
            else if (operation.Name == "BDC") {
                IReadOnlyList<string> values = ReadAlternateText(operation.Operands, propertyText);
                AlternateTextScope? scope = values.Count > 0 ? new AlternateTextScope(start, previousEnd, content.Length, values) : null;
                if (scope is not null) scopes.Add(scope);
                open.Push(scope);
            } else if (operation.Name == "EMC") {
                if (open.Count > 0 && open.Pop() is AlternateTextScope closed) closed.End = operation.OperatorOffset;
            } else if (IsPaintingOperator(operation.Name)) {
                foreach (AlternateTextScope? scope in open) if (scope is not null) scope.Paints = true;
            }
        }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
        return scopes;
    }

    private static IReadOnlyList<string> ReadAlternateText(IReadOnlyList<object> operands, PropertyAlternateText propertyText) {
        // Optional-content sections decide visibility; their property list is never alternate text.
        if (operands.Count < 2 || operands[operands.Count - 2] is "OC") return Array.Empty<string>();
        if (operands[operands.Count - 1] is string name) return propertyText(name);
        var values = new List<string>();
        if (operands[operands.Count - 1] is PdfContentDictionary dictionary) {
            foreach (string key in AlternateTextKeys) {
                if (dictionary.Items.TryGetValue(key, out object? value) && value is byte[] bytes) values.Add(PdfTextString.Decode(bytes));
            }
        }
        return values;
    }

    // A section naming a property list holds its alternate text in /Properties, which other sections can share.
    private static PropertyAlternateText ReadPropertyAlternateText(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? resources) => name => {
        if (GetPropertyLists(objects, resources) is not PdfDictionary properties || !properties.Items.TryGetValue(name, out PdfObject? entry) ||
            ResolveDictionary(objects, entry) is not PdfDictionary list) return Array.Empty<string>();
        return AlternateTextKeys.Select(key => list.Items.TryGetValue(key, out PdfObject? value) ? PdfObjectLookup.Resolve(objects, value) as PdfStringObj : null)
            .OfType<PdfStringObj>().Select(static text => PdfTextString.Decode(text.RawBytes)).ToArray();
    };

    private static PdfDictionary? GetPropertyLists(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? resources) =>
        resources is not null && resources.Items.TryGetValue("Properties", out PdfObject? value) ? ResolveDictionary(objects, value) : null;

    private static bool IsPaintingOperator(string name) =>
        name is "Tj" or "TJ" or "'" or "\"" or "Do" or "sh" or "BI" or "EI" || IsPathPaintOperator(name);

    /// <summary>
    /// Alternate text repeats what its section paints, so a section losing content loses its property list too.
    /// The section stays open so its matching EMC still balances.
    /// </summary>
    private static TextObjectRewrite[] ClearAlternateTextAround(string content, IEnumerable<TextObjectRewrite> rewrites, IEnumerable<(int Start, int End)> changes, PropertyAlternateText propertyText, PdfReadLimits limits) {
        List<TextObjectRewrite> result = rewrites.ToList();
        (int Start, int End)[] changed = changes.ToArray();
        if (changed.Length > 0) {
            // Last first, so an edit can't shift the operator an earlier section is counted against.
            foreach (AlternateTextScope scope in Enumerable.Reverse(FindAlternateTextScopes(content, propertyText, limits))) {
                if (!changed.Any(change => change.Start < scope.End && change.End > scope.OperationEnd)) continue;
                int owner = result.FindIndex(rewrite => rewrite.Index <= scope.Start && scope.OperationEnd <= rewrite.Index + rewrite.Length);
                if (owner < 0) {
                    result.Add(new TextObjectRewrite(scope.Start, scope.OperationEnd - scope.Start, PlainMarkedContent));
                } else {
                    bool changedElsewhere = changed.Any(change => change.Start != result[owner].Index && change.Start < scope.End && change.End > scope.OperationEnd);
                    result[owner] = new TextObjectRewrite(result[owner].Index, result[owner].Length, ClearOperationInRewrite(content, result[owner], scope, changedElsewhere));
                }
            }
        }
        return result.OrderBy(static rewrite => rewrite.Index).ToArray();
    }

    // A rewritten text object keeps its other operators verbatim and in order, so the nth copy of this one is still
    // the nth. A section inside the object can repeat exactly, and only one that really changed loses its list.
    private static string ClearOperationInRewrite(string content, TextObjectRewrite rewrite, AlternateTextScope scope, bool changedElsewhere) {
        string operation = content.Substring(scope.Start, scope.OperationEnd - scope.Start);
        string replacement = rewrite.Replacement;
        int position = replacement.IndexOf(operation, StringComparison.Ordinal);
        for (int source = content.IndexOf(operation, rewrite.Index, StringComparison.Ordinal);
            source >= 0 && source < scope.Start && position >= 0;
            source = content.IndexOf(operation, source + operation.Length, StringComparison.Ordinal)) {
            position = replacement.IndexOf(operation, position + operation.Length, StringComparison.Ordinal);
        }
        if (position < 0) return replacement;
        int bodyLength = Math.Min(scope.End, rewrite.Index + rewrite.Length) - scope.OperationEnd;
        int body = position + operation.Length;
        if (!changedElsewhere && body + bodyLength <= replacement.Length && string.CompareOrdinal(replacement, body, content, scope.OperationEnd, bodyLength) == 0) return replacement;
        return replacement.Remove(position, operation.Length).Insert(position, PlainMarkedContent);
    }

    private static TextObjectRewrite[] ClearAlternateTextAroundRewrites(string content, IReadOnlyList<TextObjectRewrite> rewrites, PropertyAlternateText propertyText, PdfReadLimits limits) =>
        ClearAlternateTextAround(content, rewrites, rewrites
            .Where(rewrite => rewrite.Length != rewrite.Replacement.Length || string.CompareOrdinal(content, rewrite.Index, rewrite.Replacement, 0, rewrite.Length) != 0)
            .Select(static rewrite => (rewrite.Index, rewrite.Index + rewrite.Length)), propertyText, limits);

    // Nothing in the owning stream changed: the invoked image or form did.
    private static string ClearAlternateTextAroundInvocations(string content, HashSet<string> changedNames, PropertyAlternateText propertyText, PdfReadLimits limits) {
        if (changedNames.Count == 0) return content;
        var invocations = new List<(int Start, int End)>();
        PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations, operation => {
            if (operation.Name == "Do" && operation.Operands.Count > 0 && operation.Operands[operation.Operands.Count - 1] is string name && changedNames.Contains(name))
                invocations.Add((operation.OperatorOffset, operation.OperatorOffset + 2));
        }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
        return RewriteTextObjectSpans(content, 0, ClearAlternateTextAround(content, Array.Empty<TextObjectRewrite>(), invocations, propertyText, limits));
    }

    // Clearing a section stops it naming its property list, but the value stays in the file while /Properties lists
    // it. A list other sections still name stays with them. Unreadable content could name any entry, so nothing goes.
    private static bool RemoveUnnamedAlternateTextProperties(Dictionary<int, PdfIndirectObject> objects, PdfReadLimits limits, HashSet<PdfStream> sourceStreamIdentities) {
        var named = new Dictionary<PdfDictionary, HashSet<string>>();
        var inherited = new HashSet<string>(StringComparer.Ordinal);
        foreach ((PdfDictionary? resources, IReadOnlyList<PdfStream> streams) in EnumerateResourceOwners(objects)) {
            PdfDictionary? lists = GetPropertyLists(objects, resources);
            HashSet<string> names = resources is null ? inherited : lists is null ? new HashSet<string>(StringComparer.Ordinal)
                : named.TryGetValue(lists, out HashSet<string>? existing) ? existing : named[lists] = new HashSet<string>(StringComparer.Ordinal);
            foreach (PdfStream stream in streams) {
                if (stream.DecodingFailed || !StreamDecoder.TryDecode(stream, GetMutationDecodeLimit(stream, limits, sourceStreamIdentities), out byte[] decoded, objects)) return false;
                names.UnionWith(FindNamedPropertyLists(PdfEncoding.Latin1GetString(decoded), limits));
            }
        }
        bool changed = false;
        foreach (KeyValuePair<PdfDictionary, HashSet<string>> list in named) {
            PropertyAlternateText propertyText = ReadPropertyAlternateText(objects, new PdfDictionary { Items = { ["Properties"] = list.Key } });
            foreach (string name in list.Key.Items.Keys.ToArray()) {
                if (!list.Value.Contains(name) && !inherited.Contains(name) && propertyText(name).Count > 0) changed = list.Key.Items.Remove(name) || changed;
            }
        }
        return changed;
    }

    private static HashSet<string> FindNamedPropertyLists(string content, PdfReadLimits limits) {
        var names = new HashSet<string>(StringComparer.Ordinal);
        if (content.IndexOf("BDC", StringComparison.Ordinal) < 0) return names;
        PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations, operation => {
            if (operation.Name == "BDC" && operation.Operands.Count > 1 && operation.Operands[operation.Operands.Count - 1] is string name) names.Add(name);
        }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
        return names;
    }

    // Resources are commonly shared between pages. A page-scoped search leaves the other pages' sections alone,
    // so a searched list this page no longer names leaves this page's resources instead of being edited in place.
    private static void DetachUnnamedSearchedProperties(Dictionary<int, PdfIndirectObject> objects, int[] pages, PdfRedactionSearchCriteria[] criteria, PdfReadLimits limits, HashSet<PdfStream> sourceStreamIdentities) {
        for (int index = 0; index < pages.Length; index++) {
            if (!criteria.Any(criterion => criterion.PageNumbers.Contains(index + 1)) || !objects.TryGetValue(pages[index], out PdfIndirectObject? item) ||
                item.Value is not PdfDictionary page || !page.Items.TryGetValue("Contents", out PdfObject? contents) ||
                GetInheritedDictionary(objects, page, "Resources") is not PdfDictionary resources || GetPropertyLists(objects, resources) is not PdfDictionary lists) continue;
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (PdfReference reference in EnumerateContentStreamReferences(objects, contents)) {
                if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) && indirect.Value is PdfStream stream && !stream.DecodingFailed &&
                    StreamDecoder.TryDecode(stream, GetMutationDecodeLimit(stream, limits, sourceStreamIdentities), out byte[] decoded, objects)) names.UnionWith(FindNamedPropertyLists(PdfEncoding.Latin1GetString(decoded), limits));
            }
            string[] detached = lists.Items.Where(entry => !names.Contains(entry.Key) && HoldsSearchedAlternateText(objects, PdfObjectLookup.Resolve(objects, entry.Value) ?? entry.Value, criteria)).Select(static entry => entry.Key).ToArray();
            if (detached.Length == 0) continue;
            PdfDictionary ownResources = CloneDictionary(resources), ownLists = CloneDictionary(lists);
            foreach (string name in detached) ownLists.Items.Remove(name);
            ownResources.Items["Properties"] = ownLists;
            page.Items["Resources"] = ownResources;
        }
    }

    /// <summary>Finds alternate text the rewrite should have cleared: a section left painting nothing, or one repeating reviewed text.</summary>
    internal static IReadOnlyList<PdfRedactionVerificationIssue> FindAlternateTextAroundRemovedContent(
        byte[] pdf, PdfRedactionPlan plan, PdfLoadOptions? readOptions, System.Threading.CancellationToken cancellationToken) {
        var issues = new List<PdfRedactionVerificationIssue>();
        var pages = new HashSet<int>(plan.Matches.Select(static match => match.PageNumber));
        if (pages.Count == 0) return issues;
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions, cancellationToken);
        Dictionary<int, PdfIndirectObject> objects = document.Objects;
        PdfReadLimits limits = document.ReadOptions.Limits;
        var reviewed = new HashSet<string>(plan.ReviewedTextObjectScopes.SelectMany(static scopes => scopes).SelectMany(static scope => scope.Spans)
            .Select(static span => span.SourceActualText).OfType<string>(), StringComparer.Ordinal);
        foreach (int pageNumber in pages) {
            cancellationToken.ThrowIfCancellationRequested();
            if (pageNumber < 1 || pageNumber > document.Pages.Count ||
                !objects.TryGetValue(document.Pages[pageNumber - 1].ObjectNumber, out PdfIndirectObject? pageObject) || pageObject.Value is not PdfDictionary page ||
                !page.Items.TryGetValue("Contents", out PdfObject? contents)) continue;
            var content = new StringBuilder();
            foreach (PdfReference reference in EnumerateContentReferences(objects, contents)) {
                if (PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) && indirect.Value is PdfStream stream &&
                    TryDecodeContent(stream, objects, limits, out string segment)) content.Append(segment).Append('\n');
            }
            Check(content.ToString(), GetInheritedDictionary(objects, page, "Resources"), "page " + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        foreach (PdfIndirectObject item in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (item.Value is PdfStream form && string.Equals(form.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal) &&
                TryDecodeContent(form, objects, limits, out string formContent)) Check(formContent, ResolveDictionary(objects, form.Dictionary.Items.TryGetValue("Resources", out PdfObject? own) ? own : null),
                "object " + item.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        return issues;

        void Check(string content, PdfDictionary? resources, string owner) {
            if (FindAlternateTextScopes(content, ReadPropertyAlternateText(objects, resources), limits).Any(scope => !scope.Paints || scope.Values.Any(reviewed.Contains)))
                issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "AlternateText", "Marked content on " + owner + " keeps alternate text for removed content."));
        }
    }
}
