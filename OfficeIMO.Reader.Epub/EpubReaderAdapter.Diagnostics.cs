using OfficeIMO.Epub;
using OfficeIMO.Reader;

namespace OfficeIMO.Reader.Epub;

internal static partial class EpubReaderAdapter {
    private static IReadOnlyList<OfficeDocumentDiagnostic> BuildEpubDiagnostics(
        EpubDocument document,
        string sourcePath) {
        if (document.Diagnostics.Count == 0) return Array.Empty<OfficeDocumentDiagnostic>();

        var diagnostics = new List<OfficeDocumentDiagnostic>(document.Diagnostics.Count);
        foreach (EpubDiagnostic diagnostic in document.Diagnostics) {
            diagnostics.Add(new OfficeDocumentDiagnostic {
                Severity = MapEpubDiagnosticSeverity(diagnostic.Severity),
                Category = MapEpubDiagnosticCategory(diagnostic.Code),
                Code = diagnostic.Code,
                Message = diagnostic.Message,
                Source = "OfficeIMO.Epub",
                IsRecoverable = diagnostic.Severity != EpubDiagnosticSeverity.Error,
                Location = new ReaderLocation {
                    Path = string.IsNullOrWhiteSpace(diagnostic.Path)
                        ? sourcePath
                        : BuildEpubLocationPath(sourcePath, diagnostic.Path!)
                }
            });
        }
        return diagnostics;
    }

    private static OfficeDocumentDiagnosticSeverity MapEpubDiagnosticSeverity(EpubDiagnosticSeverity severity) {
        return severity switch {
            EpubDiagnosticSeverity.Info => OfficeDocumentDiagnosticSeverity.Information,
            EpubDiagnosticSeverity.Error => OfficeDocumentDiagnosticSeverity.Error,
            _ => OfficeDocumentDiagnosticSeverity.Warning
        };
    }

    private static OfficeDocumentDiagnosticCategory MapEpubDiagnosticCategory(string code) {
        if (code.IndexOf("encrypt", StringComparison.Ordinal) >= 0) {
            return OfficeDocumentDiagnosticCategory.Security;
        }
        if (code.IndexOf("limit", StringComparison.Ordinal) >= 0) {
            return OfficeDocumentDiagnosticCategory.Limit;
        }
        if (code.IndexOf("layout", StringComparison.Ordinal) >= 0) {
            return OfficeDocumentDiagnosticCategory.Content;
        }
        return OfficeDocumentDiagnosticCategory.Parsing;
    }
}
