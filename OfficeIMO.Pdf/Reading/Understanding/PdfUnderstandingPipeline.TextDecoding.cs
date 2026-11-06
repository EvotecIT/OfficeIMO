using System.Globalization;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfUnderstandingPipeline {
    // Retain raw parser/paint evidence for inspection and fail-closed redaction.
    // Only the semantic input loses runs that cannot supply Unicode text.
    private static IReadOnlyList<PdfTextSpan> FilterUndecodableRuns(
        PdfUnderstandingPageContext context, IReadOnlyList<PdfTextSpan> runs) {
        List<PdfTextSpan>? retained = null;
        List<PdfFontInspectionDiagnostic>? diagnostics = null;
        HashSet<(PdfFontInspectionDiagnosticCode Code, string Font, int? Stream)>? seen = null;
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
                diagnostics = new List<PdfFontInspectionDiagnostic>();
                seen = new HashSet<(PdfFontInspectionDiagnosticCode, string, int?)>();
            }
            if (!seen!.Add((code, run.FontResource, run.ContentStreamObjectNumber))) continue;
            context.ConsumeWork(2L * run.FontResource.Length + 80L);
            string path = "Page " + context.PageNumber.ToString(CultureInfo.InvariantCulture);
            if (run.ContentStreamObjectNumber is int stream)
                path += "/ContentStream/" + stream.ToString(CultureInfo.InvariantCulture);
            path += "/Font/" + run.FontResource;
            string reason = code == PdfFontInspectionDiagnosticCode.MissingFontResource
                ? "the font resource could not be resolved"
                : "its character codes could not be mapped to Unicode";
            diagnostics!.Add(new PdfFontInspectionDiagnostic(code,
                "Text using font /" + run.FontResource + " was omitted because " + reason + ".",
                context.PageNumber, path));
        }
        if (retained is null) return runs;
        context.TextDecodingDiagnostics = diagnostics!.AsReadOnly();
        return retained.AsReadOnly();
    }
}
