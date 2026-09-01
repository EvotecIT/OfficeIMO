namespace OfficeIMO.Pdf;

/// <summary>Edits the catalog JavaScript name tree through the shared object-graph rewrite owner.</summary>
internal static class PdfJavaScriptEditor {
    internal static PdfJavaScriptEditResult Edit(
        byte[] pdf,
        Action<PdfJavaScriptEditSession> edit,
        PdfLoadOptions? readOptions = null) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(edit, nameof(edit));

        PdfReadDocument source = PdfReadDocument.Open(pdf, readOptions);
        var session = new PdfJavaScriptEditSession(source.ReadOptions.Limits);
        edit(session);
        if (session.Operations.Count == 0) {
            throw new ArgumentException("At least one document JavaScript edit command is required.", nameof(edit));
        }

        PdfMutationPlan plan = PdfMutationPlanner.RequireFullRewrite(
            pdf,
            PdfMutationOperation.ModifyJavaScript,
            readOptions);
        byte[] untouchedJavaScriptSnapshot = PdfJavaScriptNameTreeEditor.CreateUntouchedSnapshot(
            source.Objects,
            source.Security,
            session.Commands,
            source.ReadOptions.Limits);
        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(
            pdf,
            readOptions,
            outputEncryption: null,
            (objects, security) => {
                PdfJavaScriptNameTreeEditor.Rewrite(objects, security, session.Commands, source.ReadOptions.Limits);
                return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value)
                    ? security.InfoObjectNumber
                    : null;
            });

        PdfLoadOptions rewrittenReadOptions = PdfLoadOptions.WithMinimumInputBytes(source.ReadOptions, output.LongLength);
        PdfReadDocument savedDocument = PdfReadDocument.Open(output, rewrittenReadOptions);
        IReadOnlyList<PdfJavaScript> saved = savedDocument.JavaScripts;
        ValidateReadback(session.Commands, saved);
        ValidateUntouchedJavaScript(
            untouchedJavaScriptSnapshot,
            PdfJavaScriptNameTreeEditor.CreateUntouchedSnapshot(
                savedDocument.Objects,
                savedDocument.Security,
                session.Commands,
                savedDocument.ReadOptions.Limits));
        ValidateOtherCatalogActions(source.CatalogActions, savedDocument.CatalogActions);
        var preservationOptions = new PdfRewritePreservationOptions {
            OriginalReadOptions = source.ReadOptions,
            RewrittenReadOptions = rewrittenReadOptions,
            PreserveCatalogActions = false,
            PreserveRevisionStructure = false
        };
        PdfRewritePreservationReport preservation = PdfRewritePreservation.AssertPreserved(
            pdf,
            output,
            preservationOptions);
        return new PdfJavaScriptEditResult(output, plan, preservation, saved, session.Operations, rewrittenReadOptions);
    }

    internal static PdfJavaScriptEditResult AddOrReplace(
        byte[] pdf,
        string name,
        string script,
        PdfLoadOptions? readOptions = null) =>
        Edit(pdf, session => session.AddOrReplace(name, script), readOptions);

    internal static PdfJavaScriptEditResult Remove(
        byte[] pdf,
        string name,
        PdfLoadOptions? readOptions = null) =>
        Edit(pdf, session => session.Remove(name), readOptions);

    internal static PdfJavaScriptEditResult Clear(byte[] pdf, PdfLoadOptions? readOptions = null) =>
        Edit(pdf, static session => session.Clear(), readOptions);

    private static void ValidateReadback(
        IReadOnlyList<PdfJavaScriptEditSession.EditCommand> commands,
        IReadOnlyList<PdfJavaScript> actual) {
        var expected = new Dictionary<string, string?>(StringComparer.Ordinal);
        bool cleared = false;
        for (int i = 0; i < commands.Count; i++) {
            PdfJavaScriptEditSession.EditCommand command = commands[i];
            if (command.Kind == PdfJavaScriptEditSession.EditKind.Clear) {
                expected.Clear(); cleared = true;
            } else {
                expected[command.Name!] = command.Kind == PdfJavaScriptEditSession.EditKind.Remove ? null : command.Script;
            }
        }
        foreach (KeyValuePair<string, string?> expectation in expected) {
            PdfJavaScript[] matches = actual.Where(script => string.Equals(script.Name, expectation.Key, StringComparison.Ordinal)).ToArray();
            if (expectation.Value is null && matches.Length == 0) continue;
            if (expectation.Value is not null && matches.Length == 1 && string.Equals(matches[0].Script, expectation.Value, StringComparison.Ordinal)) continue;
            throw new InvalidOperationException("PDF document JavaScript post-save validation found a name or source mismatch; the artifact was not returned.");
        }
        if (cleared && actual.Count != expected.Count(static item => item.Value is not null)) {
            throw new InvalidOperationException("PDF document JavaScript post-save validation found an unexpected script after clearing the name tree; the artifact was not returned.");
        }
    }

    private static void ValidateOtherCatalogActions(
        IReadOnlyList<PdfCatalogAction> expected,
        IReadOnlyList<PdfCatalogAction> actual) {
        string[] expectedValues = expected
            .Where(static action => !string.Equals(action.Source, "Names/JavaScript", StringComparison.Ordinal))
            .Select(GetCatalogActionIdentity)
            .ToArray();
        string[] actualValues = actual
            .Where(static action => !string.Equals(action.Source, "Names/JavaScript", StringComparison.Ordinal))
            .Select(GetCatalogActionIdentity)
            .ToArray();
        if (!expectedValues.SequenceEqual(actualValues, StringComparer.Ordinal)) {
            throw new InvalidOperationException("PDF document JavaScript post-save validation found an unrelated catalog-action change; the artifact was not returned.");
        }
    }

    private static void ValidateUntouchedJavaScript(byte[] expected, byte[] actual) {
        if (!expected.SequenceEqual(actual)) {
            throw new InvalidOperationException("PDF document JavaScript post-save validation found an untouched name-tree entry or action-graph change; the artifact was not returned.");
        }
    }

    private static string GetCatalogActionIdentity(PdfCatalogAction action) =>
        action.Name + "\u001f" +
        action.ActionType + "\u001f" +
        action.Source + "\u001f" +
        (action.TriggerName ?? string.Empty) + "\u001f" +
        (action.ActionPath ?? string.Empty) + "\u001f" +
        (action.IsChainedAction ? "1" : "0");
}
