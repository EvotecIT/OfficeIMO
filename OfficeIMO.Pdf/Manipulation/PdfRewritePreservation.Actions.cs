namespace OfficeIMO.Pdf;

public static partial class PdfRewritePreservation {
    private static void CompareViewerActionState(List<PdfRewritePreservationIssue> issues, PdfDocumentInfo original, PdfDocumentInfo rewritten, PdfRewritePreservationOptions options) {
        CompareOpenAction(issues, original.OpenAction, rewritten.OpenAction, options);
        CompareViewerPreferences(issues, original.ViewerPreferences, rewritten.ViewerPreferences, options);
        CompareCatalogActions(issues, original.CatalogActions, rewritten.CatalogActions, options);
        ComparePageActions(issues, original.Pages, rewritten.Pages, options);
    }

    private static void CompareOpenAction(List<PdfRewritePreservationIssue> issues, PdfDocumentOpenAction? original, PdfDocumentOpenAction? rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveOpenAction) {
            return;
        }

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
        if (!options.PreserveCatalogActions || original.Count != rewritten.Count) {
            return;
        }

        for (int i = 0; i < original.Count; i++) {
            PdfCatalogAction before = original[i];
            PdfCatalogAction after = rewritten[i];
            string prefix = "CatalogActions[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

            CompareString(issues, prefix + ".Name", before.Name, after.Name);
            CompareString(issues, prefix + ".ActionType", before.ActionType, after.ActionType);
            CompareString(issues, prefix + ".Source", before.Source, after.Source);
            CompareString(issues, prefix + ".TriggerName", before.TriggerName, after.TriggerName);
        }
    }

    private static void ComparePageActions(List<PdfRewritePreservationIssue> issues, IReadOnlyList<PdfPageInfo> originalPages, IReadOnlyList<PdfPageInfo> rewrittenPages, PdfRewritePreservationOptions options) {
        if (!options.PreservePageActions || originalPages.Count != rewrittenPages.Count) {
            return;
        }

        for (int i = 0; i < originalPages.Count; i++) {
            IReadOnlyList<PdfPageAction> original = originalPages[i].PageActions;
            IReadOnlyList<PdfPageAction> rewritten = rewrittenPages[i].PageActions;
            if (original.Count != rewritten.Count) {
                continue;
            }

            for (int j = 0; j < original.Count; j++) {
                PdfPageAction before = original[j];
                PdfPageAction after = rewritten[j];
                string prefix = "PageActions[" + originalPages[i].PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + "," + j.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

                CompareNullableInt(issues, prefix + ".PageNumber", before.PageNumber, after.PageNumber);
                CompareString(issues, prefix + ".TriggerName", before.TriggerName, after.TriggerName);
                CompareString(issues, prefix + ".ActionType", before.ActionType, after.ActionType);
                CompareString(issues, prefix + ".ActionPath", before.ActionPath, after.ActionPath);
            }
        }
    }

    private static void CompareNullablePresence(List<PdfRewritePreservationIssue> issues, string feature, bool expectedPresent, bool actualPresent) {
        if (expectedPresent == actualPresent) {
            return;
        }

        issues.Add(CreateIssue(feature, expectedPresent ? "present" : "missing", actualPresent ? "present" : "missing"));
    }
}
