namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private const string PlainMarkedContent = "/Span BMC";

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
    private static List<AlternateTextScope> FindAlternateTextScopes(string content, PdfReadLimits limits) {
        var scopes = new List<AlternateTextScope>();
        if (content.IndexOf("BDC", StringComparison.Ordinal) < 0) return scopes;
        var open = new Stack<AlternateTextScope?>();
        int previousEnd = 0;
        PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations, operation => {
            int start = PdfContentStreamTextRewriter.FindOperationStart(content, previousEnd, operation.OperatorOffset);
            previousEnd = Math.Min(content.Length, operation.OperatorOffset + operation.Name.Length);
            if (operation.Name == "BMC") open.Push(null);
            else if (operation.Name == "BDC") {
                var values = new List<string>();
                if (operation.Operands.Count > 0 && operation.Operands[operation.Operands.Count - 1] is PdfContentDictionary dictionary) {
                    foreach (string key in AlternateTextKeys) {
                        if (dictionary.Items.TryGetValue(key, out object? value) && value is byte[] bytes) values.Add(PdfTextString.Decode(bytes));
                    }
                }
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

    private static bool IsPaintingOperator(string name) =>
        name is "Tj" or "TJ" or "'" or "\"" or "Do" or "sh" or "BI" or "EI" || IsPathPaintOperator(name);

    /// <summary>
    /// Alternate text repeats what its section paints, so a section losing content loses its property list too.
    /// The section stays open so its matching EMC still balances.
    /// </summary>
    private static TextObjectRewrite[] ClearAlternateTextAround(string content, IEnumerable<TextObjectRewrite> rewrites, IEnumerable<(int Start, int End)> changes, PdfReadLimits limits) {
        List<TextObjectRewrite> result = rewrites.ToList();
        (int Start, int End)[] changed = changes.ToArray();
        if (changed.Length > 0) {
            foreach (AlternateTextScope scope in FindAlternateTextScopes(content, limits)) {
                if (!changed.Any(change => change.Start < scope.End && change.End > scope.OperationEnd)) continue;
                int owner = result.FindIndex(rewrite => rewrite.Index <= scope.Start && scope.OperationEnd <= rewrite.Index + rewrite.Length);
                if (owner < 0) {
                    result.Add(new TextObjectRewrite(scope.Start, scope.OperationEnd - scope.Start, PlainMarkedContent));
                } else {
                    // A rewritten text object keeps its other operators verbatim, this one included.
                    string operation = content.Substring(scope.Start, scope.OperationEnd - scope.Start);
                    result[owner] = new TextObjectRewrite(result[owner].Index, result[owner].Length, result[owner].Replacement.Replace(operation, PlainMarkedContent));
                }
            }
        }
        return result.OrderBy(static rewrite => rewrite.Index).ToArray();
    }

    private static TextObjectRewrite[] ClearAlternateTextAroundRewrites(string content, IReadOnlyList<TextObjectRewrite> rewrites, PdfReadLimits limits) =>
        ClearAlternateTextAround(content, rewrites, rewrites
            .Where(rewrite => rewrite.Length != rewrite.Replacement.Length || string.CompareOrdinal(content, rewrite.Index, rewrite.Replacement, 0, rewrite.Length) != 0)
            .Select(static rewrite => (rewrite.Index, rewrite.Index + rewrite.Length)), limits);

    // Nothing in the owning stream changed: the invoked image or form did.
    private static string ClearAlternateTextAroundInvocations(string content, HashSet<string> changedNames, PdfReadLimits limits) {
        if (changedNames.Count == 0) return content;
        var invocations = new List<(int Start, int End)>();
        PdfContentStreamInterpreter.Interpret(content, limits.MaxContentOperations, operation => {
            if (operation.Name == "Do" && operation.Operands.Count > 0 && operation.Operands[operation.Operands.Count - 1] is string name && changedNames.Contains(name))
                invocations.Add((operation.OperatorOffset, operation.OperatorOffset + 2));
        }, maxNestingDepth: limits.MaxContentNestingDepth, maxOperands: limits.MaxContentOperands);
        return RewriteTextObjectSpans(content, 0, ClearAlternateTextAround(content, Array.Empty<TextObjectRewrite>(), invocations, limits));
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
            Check(content.ToString(), "page " + pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        foreach (PdfIndirectObject item in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (item.Value is PdfStream form && string.Equals(form.Dictionary.Get<PdfName>("Subtype")?.Name, "Form", StringComparison.Ordinal) &&
                TryDecodeContent(form, objects, limits, out string formContent)) Check(formContent, "object " + item.ObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        return issues;

        void Check(string content, string owner) {
            if (FindAlternateTextScopes(content, limits).Any(scope => !scope.Paints || scope.Values.Any(reviewed.Contains)))
                issues.Add(new PdfRedactionVerificationIssue("RedactedContentOutsidePageContent", "AlternateText", "Marked content on " + owner + " keeps alternate text for removed content."));
        }
    }
}
