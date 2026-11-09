using System.Text;

namespace OfficeIMO;

/// <summary>Read-only validation findings and the native evidence that produced them.</summary>
public sealed class OfficeValidationReport {
    /// <summary>Creates a report with detached collections of diagnostics and native reports.</summary>
    public OfficeValidationReport(IEnumerable<OfficeDiagnostic> diagnostics, IEnumerable<object>? nativeReports = null) {
        if (diagnostics == null) throw new ArgumentNullException(nameof(diagnostics));
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        NativeReports = Array.AsReadOnly((nativeReports ?? Array.Empty<object>()).ToArray());
    }

    /// <summary>True when no performed check reported an error. Not-applicable checks are explicitly identified in Diagnostics.</summary>
    public bool IsValid => !Diagnostics.Any(item => item.Severity == OfficeDiagnosticSeverity.Error);
    /// <summary>Findings, including explicit Validation.NotApplicable observations for checks unavailable on this input.</summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics { get; }
    /// <summary>Native validation evidence, including package inventories, feature reports and PDF proof requirements.</summary>
    public IReadOnlyList<object> NativeReports { get; }

    /// <summary>Formats the validation result and its diagnostics for a Markdown report.</summary>
    public string ToMarkdown() {
        var text = new StringBuilder();
        text.AppendLine("# Document validation");
        text.AppendLine();
        text.AppendLine(IsValid ? "Valid: no errors were reported by the performed checks." : "Invalid: one or more checks reported errors.");
        if (Diagnostics.Count != 0) {
            text.AppendLine();
            text.AppendLine("| Severity | Code | Location | Diagnostic |");
            text.AppendLine("| --- | --- | --- | --- |");
            foreach (var finding in Diagnostics) text.Append("| ").Append(finding.Severity).Append(" | ")
                .Append(MarkdownCell(finding.Code)).Append(" | ").Append(MarkdownCell(finding.Location)).Append(" | ")
                .Append(MarkdownCell(finding.Message)).AppendLine(" |");
        }
        return text.ToString();
    }

    /// <summary>Throws InvalidOperationException with the failed check diagnostics when IsValid is false.</summary>
    public void ThrowIfFailed() {
        if (!IsValid) throw new InvalidOperationException(string.Join(Environment.NewLine,
            Diagnostics.Where(finding => finding.Severity == OfficeDiagnosticSeverity.Error)
                .Select(finding => finding.Code + ": " + finding.Message)));
    }

    private static string MarkdownCell(string? text) => (text ?? string.Empty).Replace("&", "&amp;")
        .Replace("<", "&lt;").Replace(">", "&gt;").Replace("|", "&#124;").Replace("\r\n", "<br>")
        .Replace("\n", "<br>").Replace("\r", "<br>");
}
