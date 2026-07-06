namespace OfficeIMO.Pdf;

public static partial class PdfRewritePreservation {
    private static void CompareSourceStructure(List<PdfRewritePreservationIssue> issues, PdfDocumentInfo original, PdfDocumentInfo rewritten, PdfRewritePreservationOptions options) {
        if (options.PreserveDocumentVersionState) {
            CompareString(issues, "SourceStructure.HeaderVersion", original.HeaderVersion, rewritten.HeaderVersion);
            CompareString(issues, "SourceStructure.CatalogVersion", original.CatalogVersion, rewritten.CatalogVersion);
            CompareString(issues, "SourceStructure.EffectiveVersion", original.EffectiveVersion, rewritten.EffectiveVersion);
        }

        if (!options.PreserveRevisionStructure) {
            return;
        }

        CompareMinimumCount(issues, "SourceStructure.StartXrefCount", original.Security.StartXrefCount, rewritten.Security.StartXrefCount, original.Security.StartXrefCount > 1);
        CompareMinimumCount(issues, "SourceStructure.RevisionCount", original.Security.RevisionCount, rewritten.Security.RevisionCount, original.Security.RevisionCount > 1);
        CompareBooleanMarker(issues, "SourceStructure.PreviousRevision", original.Security.HasPreviousRevision, rewritten.Security.HasPreviousRevision, true);
        CompareBooleanMarker(issues, "SourceStructure.IncrementalUpdates", original.Security.HasIncrementalUpdates, rewritten.Security.HasIncrementalUpdates, true);
        CompareBooleanMarker(issues, "SourceStructure.XrefStreams", original.Security.HasXrefStreams, rewritten.Security.HasXrefStreams, true);
        CompareBooleanMarker(issues, "SourceStructure.ObjectStreams", original.Security.HasObjectStreams, rewritten.Security.HasObjectStreams, true);
    }
}
