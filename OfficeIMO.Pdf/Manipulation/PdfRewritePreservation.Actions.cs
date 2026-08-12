namespace OfficeIMO.Pdf;

public static partial class PdfRewritePreservation {
    private static void CompareViewerActionState(List<PdfRewritePreservationIssue> issues, PdfDocumentInfo original, PdfDocumentInfo rewritten, PdfRewritePreservationOptions options) {
        CompareOpenAction(issues, original.OpenAction, rewritten.OpenAction, options);
        CompareViewerPreferences(issues, original.ViewerPreferences, rewritten.ViewerPreferences, options);
        CompareCatalogActions(issues, original.CatalogActions, rewritten.CatalogActions, options);
        ComparePageActions(issues, original.Pages, rewritten.Pages, options);
        CompareFormWidgetActions(issues, original.FormFields, rewritten.FormFields, options);
    }

    private static void CompareOpenAction(List<PdfRewritePreservationIssue> issues, PdfDocumentOpenAction? original, PdfDocumentOpenAction? rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveOpenAction) {
            return;
        }

        original = IsPreservedActionType(options, original?.ActionType) ? original : null;
        rewritten = IsPreservedActionType(options, rewritten?.ActionType) ? rewritten : null;

        if (original is null || rewritten is null) {
            CompareNullablePresence(issues, "OpenAction", original is not null, rewritten is not null);
            return;
        }

        CompareString(issues, "OpenAction.ActionType", original.ActionType, rewritten.ActionType);
        CompareNullableInt(issues, "OpenAction.PageNumber", original.PageNumber, rewritten.PageNumber);
        CompareNullableDestinationMode(issues, "OpenAction.DestinationMode", original.DestinationMode, rewritten.DestinationMode);
        CompareNullableDouble(issues, "OpenAction.DestinationTop", original.DestinationTop, rewritten.DestinationTop);
        CompareNullableDouble(issues, "OpenAction.DestinationLeft", original.DestinationLeft, rewritten.DestinationLeft);
        CompareNullableDouble(issues, "OpenAction.DestinationBottom", original.DestinationBottom, rewritten.DestinationBottom);
        CompareNullableDouble(issues, "OpenAction.DestinationRight", original.DestinationRight, rewritten.DestinationRight);
        CompareNullableDouble(issues, "OpenAction.DestinationZoom", original.DestinationZoom, rewritten.DestinationZoom);
    }

    private static void CompareViewerPreferences(List<PdfRewritePreservationIssue> issues, PdfViewerPreferences? original, PdfViewerPreferences? rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveViewerPreferences) {
            return;
        }

        if (original is null || rewritten is null) {
            CompareNullablePresence(issues, "ViewerPreferences", original is not null, rewritten is not null);
            return;
        }

        CompareStringDictionary(issues, "ViewerPreferences.Values", original.Values, rewritten.Values);
    }

    private static void CompareCatalogActions(List<PdfRewritePreservationIssue> issues, IReadOnlyList<PdfCatalogAction> original, IReadOnlyList<PdfCatalogAction> rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveCatalogActions) {
            return;
        }
        PdfCatalogAction[] expected = FilterPreservedActions(original, options);
        PdfCatalogAction[] actual = FilterPreservedActions(rewritten, options);
        if (expected.Length != actual.Length) {
            issues.Add(CreateIssue("CatalogActions.Count", expected.Length.ToString(System.Globalization.CultureInfo.InvariantCulture), actual.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            return;
        }

        for (int i = 0; i < expected.Length; i++) {
            PdfCatalogAction before = expected[i];
            PdfCatalogAction after = actual[i];
            string prefix = "CatalogActions[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

            CompareString(issues, prefix + ".Name", NormalizeFilteredActionPath(before.Name, options), NormalizeFilteredActionPath(after.Name, options));
            CompareString(issues, prefix + ".ActionType", before.ActionType, after.ActionType);
            CompareString(issues, prefix + ".Source", before.Source, after.Source);
            CompareString(issues, prefix + ".TriggerName", NormalizeFilteredActionPath(before.TriggerName, options), NormalizeFilteredActionPath(after.TriggerName, options));
            CompareString(issues, prefix + ".Uri", before.Uri, after.Uri);
        }
    }

    private static void ComparePageActions(List<PdfRewritePreservationIssue> issues, IReadOnlyList<PdfPageInfo> originalPages, IReadOnlyList<PdfPageInfo> rewrittenPages, PdfRewritePreservationOptions options) {
        if (!options.PreservePageActions || originalPages.Count != rewrittenPages.Count) {
            return;
        }

        for (int i = 0; i < originalPages.Count; i++) {
            PdfPageAction[] original = FilterPreservedActions(originalPages[i].PageActions, options);
            PdfPageAction[] rewritten = FilterPreservedActions(rewrittenPages[i].PageActions, options);
            if (original.Length != rewritten.Length) {
                issues.Add(CreateIssue(
                    "PageActions[" + originalPages[i].PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + "].Count",
                    original.Length.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    rewritten.Length.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                continue;
            }

            for (int j = 0; j < original.Length; j++) {
                PdfPageAction before = original[j];
                PdfPageAction after = rewritten[j];
                string prefix = "PageActions[" + originalPages[i].PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + "," + j.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

                CompareNullableInt(issues, prefix + ".PageNumber", before.PageNumber, after.PageNumber);
                CompareString(issues, prefix + ".TriggerName", NormalizeFilteredActionPath(before.TriggerName, options), NormalizeFilteredActionPath(after.TriggerName, options));
                CompareString(issues, prefix + ".ActionType", before.ActionType, after.ActionType);
                CompareString(issues, prefix + ".ActionPath", NormalizeFilteredActionPath(before.ActionPath, options), NormalizeFilteredActionPath(after.ActionPath, options));
                CompareString(issues, prefix + ".Uri", before.Uri, after.Uri);
            }
        }
    }

    private static bool IsPreservedActionType(PdfRewritePreservationOptions options, string? actionType) =>
        actionType is null || (!options.FilterActionsByPreservedTypes && options.PreservedActionTypes.Count == 0) || options.PreservedActionTypes.Contains(actionType);

    private static PdfCatalogAction[] FilterPreservedActions(IReadOnlyList<PdfCatalogAction> actions, PdfRewritePreservationOptions options) =>
        actions.Where(action =>
            (!options.FilterActionsByPreservedTypes && options.PreservedActionTypes.Count == 0 || options.PreservedActionTypes.Contains(action.ActionType)) &&
            (action.Uri is null || !options.ExcludedActionUris.Contains(action.Uri))).ToArray();

    private static PdfPageAction[] FilterPreservedActions(IReadOnlyList<PdfPageAction> actions, PdfRewritePreservationOptions options) =>
        actions.Where(action =>
            (!options.FilterActionsByPreservedTypes && options.PreservedActionTypes.Count == 0 || options.PreservedActionTypes.Contains(action.ActionType)) &&
            (action.Uri is null || !options.ExcludedActionUris.Contains(action.Uri))).ToArray();

    private static void CompareFormWidgetActions(List<PdfRewritePreservationIssue> issues, IReadOnlyList<PdfFormField> originalFields, IReadOnlyList<PdfFormField> rewrittenFields, PdfRewritePreservationOptions options) {
        if (!options.PreserveFormWidgetActions) return;
        string[] expected = CreateFormWidgetActionInventory(originalFields, options);
        string[] actual = CreateFormWidgetActionInventory(rewrittenFields, options);
        if (!expected.SequenceEqual(actual, StringComparer.Ordinal)) {
            issues.Add(CreateIssue("FormWidgetActions", string.Join(" | ", expected), string.Join(" | ", actual)));
        }
    }

    private static string[] CreateFormWidgetActionInventory(IReadOnlyList<PdfFormField> fields, PdfRewritePreservationOptions options) {
        var inventory = new List<string>();
        for (int fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++) {
            PdfFormField field = fields[fieldIndex];
            for (int widgetIndex = 0; widgetIndex < field.Widgets.Count; widgetIndex++) {
                PdfFormWidget widget = field.Widgets[widgetIndex];
                for (int actionIndex = 0; actionIndex < widget.Actions.Count; actionIndex++) {
                    PdfFormWidgetAction action = widget.Actions[actionIndex];
                    if (options.FilterActionsByPreservedTypes && !options.PreservedActionTypes.Contains(action.ActionType)) continue;
                    inventory.Add((field.Name ?? string.Empty) + "\u001f" +
                                  widgetIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\u001f" +
                                  action.ActionType + "\u001f" + (action.JavaScript ?? string.Empty));
                }
            }
        }
        inventory.Sort(StringComparer.Ordinal);
        return inventory.ToArray();
    }

    private static string? NormalizeFilteredActionPath(string? value, PdfRewritePreservationOptions options) {
        if (!options.FilterActionsByPreservedTypes || value is null || value.Length == 0) return value;
        const string nextMarker = ".Next.";
        int markerIndex = value.IndexOf(nextMarker, StringComparison.Ordinal);
        if (markerIndex < 0) return value;

        var normalized = new System.Text.StringBuilder(value.Length);
        int sourceIndex = 0;
        while (markerIndex >= 0) {
            int indexStart = markerIndex + nextMarker.Length;
            int indexEnd = indexStart;
            while (indexEnd < value.Length && value[indexEnd] >= '0' && value[indexEnd] <= '9') indexEnd++;
            normalized.Append(value, sourceIndex, markerIndex - sourceIndex + ".Next".Length);
            sourceIndex = indexEnd;
            markerIndex = value.IndexOf(nextMarker, sourceIndex, StringComparison.Ordinal);
        }
        normalized.Append(value, sourceIndex, value.Length - sourceIndex);
        return normalized.ToString();
    }

    private static void CompareNullablePresence(List<PdfRewritePreservationIssue> issues, string feature, bool expectedPresent, bool actualPresent) {
        if (expectedPresent == actualPresent) {
            return;
        }

        issues.Add(CreateIssue(feature, expectedPresent ? "present" : "missing", actualPresent ? "present" : "missing"));
    }
}
