namespace OfficeIMO.Pdf;

internal sealed partial class PdfUnderstandingPipeline {
    // Retain raw parser/paint evidence for inspection and fail-closed redaction.
    // Only the semantic input loses runs that cannot supply Unicode text.
    private static IReadOnlyList<PdfTextSpan> FilterUndecodableRuns(
        PdfUnderstandingPageContext context, IReadOnlyList<PdfTextSpan> runs) {
        List<PdfTextSpan>? retained = null;
        PdfTextDecodingDiagnosticsBuilder? diagnostics = context.InspectionTextRecovery;
        for (int index = 0; index < runs.Count; index++) {
            context.ConsumeWork();
            PdfTextSpan run = runs[index];
            if (run.DecodingEvidence.FailureCode is not { } code) {
                retained?.Add(run);
                continue;
            }
            if (retained is null) {
                if (index > 0) context.ConsumeWork(index);
                retained = new List<PdfTextSpan>(runs.Count);
                for (int previous = 0; previous < index; previous++) retained.Add(runs[previous]);
                diagnostics ??= new PdfTextDecodingDiagnosticsBuilder(context);
            }
            diagnostics!.Add(code, run.FontResource, run.ContentStreamObjectNumber);
        }
        if (diagnostics is not null) context.TextDecodingDiagnostics = diagnostics.ToDiagnostics();
        return retained is null ? runs : retained.AsReadOnly();
    }
}
