namespace OfficeIMO.Pdf;

internal sealed class PdfTextEncodingPreflightException : ArgumentException {
    public PdfTextEncodingPreflightException(IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics)
        : base(BuildMessage(diagnostics)) {
        TextEncodingDiagnostics = diagnostics;
        Data["textEncodingDiagnostics"] = TextEncodingDiagnostics;
        Data["diagnosticsCount"] = TextEncodingDiagnostics.Count;

        if (TextEncodingDiagnostics.Count > 0) {
            PdfTextEncodingDiagnostic first = TextEncodingDiagnostics[0];
            Data["code"] = first.Code;
            Data["source"] = first.Source;
            Data["index"] = first.Index;
            Data["codePoint"] = first.CodePoint;
            Data["text"] = first.Text;
            Data["isControlCharacter"] = first.IsControlCharacter;
            if (!string.IsNullOrWhiteSpace(first.Location)) {
                Data["location"] = first.Location;
            }

            if (!string.IsNullOrWhiteSpace(first.FieldName)) {
                Data["fieldName"] = first.FieldName;
            }

            if (first.RunIndex.HasValue) {
                Data["runIndex"] = first.RunIndex.Value;
            }

            if (first.PageNumber.HasValue) {
                Data["pageNumber"] = first.PageNumber.Value;
            }

            if (first.TableRowIndex.HasValue && first.TableColumnIndex.HasValue) {
                Data["tableRowIndex"] = first.TableRowIndex.Value;
                Data["tableColumnIndex"] = first.TableColumnIndex.Value;
            }

            if (!string.IsNullOrWhiteSpace(first.Encoding)) {
                Data["encoding"] = first.Encoding;
            }

            if (!string.IsNullOrWhiteSpace(first.Remediation)) {
                Data["remediation"] = first.Remediation;
            }
        }
    }

    public IReadOnlyList<PdfTextEncodingDiagnostic> TextEncodingDiagnostics { get; }

    private static string BuildMessage(IReadOnlyList<PdfTextEncodingDiagnostic> diagnostics) {
        int count = diagnostics == null ? 0 : diagnostics.Count;
        string summary = count == 1
            ? "PDF text encoding preflight found 1 generated text issue."
            : "PDF text encoding preflight found " + count + " generated text issues.";
        if (count == 0 || diagnostics == null) {
            return summary;
        }

        return summary + " " + diagnostics[0].Message;
    }
}
