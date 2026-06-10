using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

public static partial class DocumentReaderPdfExtensions {
    private static IReadOnlyList<OfficeDocumentDiagnostic> BuildDocumentDiagnostics(
        IReadOnlyList<ReaderChunk> chunks,
        IReadOnlyList<OfficeDocumentOcrCandidate> ocrCandidates,
        PdfDocumentPreflight? preflight) {
        var diagnostics = new List<OfficeDocumentDiagnostic>();
        AddChunkWarnings(diagnostics, chunks);
        AddOcrDiagnostics(diagnostics, ocrCandidates);
        AddPdfPreflightDiagnostics(diagnostics, preflight);

        return diagnostics.Count == 0 ? Array.Empty<OfficeDocumentDiagnostic>() : diagnostics.AsReadOnly();
    }

    private static void AddChunkWarnings(List<OfficeDocumentDiagnostic> diagnostics, IReadOnlyList<ReaderChunk> chunks) {
        for (int i = 0; i < chunks.Count; i++) {
            ReaderChunk chunk = chunks[i];
            if (chunk.Warnings == null) {
                continue;
            }

            for (int warningIndex = 0; warningIndex < chunk.Warnings.Count; warningIndex++) {
                diagnostics.Add(new OfficeDocumentDiagnostic {
                    Severity = OfficeDocumentDiagnosticSeverity.Warning,
                    Code = "reader-warning",
                    Message = chunk.Warnings[warningIndex],
                    Location = chunk.Location
                });
            }
        }
    }

    private static void AddOcrDiagnostics(List<OfficeDocumentDiagnostic> diagnostics, IReadOnlyList<OfficeDocumentOcrCandidate> ocrCandidates) {
        for (int i = 0; i < ocrCandidates.Count; i++) {
            OfficeDocumentOcrCandidate candidate = ocrCandidates[i];
            diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Code = "ocr-needed",
                Message = candidate.Reason ?? "OCR should be considered for this source region.",
                Location = candidate.Location
            });
        }
    }

    private static void AddPdfPreflightDiagnostics(List<OfficeDocumentDiagnostic> diagnostics, PdfDocumentPreflight? preflight) {
        if (preflight == null) {
            return;
        }

        for (int i = 0; i < preflight.ReadBlockers.Count; i++) {
            AddDistinctDiagnostic(
                diagnostics,
                OfficeDocumentDiagnosticSeverity.Error,
                "pdf-read-blocker",
                preflight.ReadBlockers[i].Message);
        }

        for (int i = 0; i < preflight.RewriteBlockers.Count; i++) {
            AddDistinctDiagnostic(
                diagnostics,
                OfficeDocumentDiagnosticSeverity.Warning,
                "pdf-rewrite-blocker",
                preflight.RewriteBlockers[i].Message);
        }
    }

    private static void AddDistinctDiagnostic(List<OfficeDocumentDiagnostic> diagnostics, OfficeDocumentDiagnosticSeverity severity, string code, string message) {
        if (string.IsNullOrWhiteSpace(message)) {
            return;
        }

        for (int i = 0; i < diagnostics.Count; i++) {
            OfficeDocumentDiagnostic diagnostic = diagnostics[i];
            if (diagnostic.Severity == severity &&
                string.Equals(diagnostic.Code, code, StringComparison.Ordinal) &&
                string.Equals(diagnostic.Message, message, StringComparison.Ordinal)) {
                return;
            }
        }

        diagnostics.Add(new OfficeDocumentDiagnostic {
            Severity = severity,
            Code = code,
            Message = message
        });
    }
}
