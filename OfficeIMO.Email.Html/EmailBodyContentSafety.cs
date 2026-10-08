using OfficeIMO.ContentSafety;

namespace OfficeIMO.Email;

/// <summary>Advisory evidence for one selected email body. This is not a malware or agent-safety verdict.</summary>
public sealed class EmailBodyContentSafetyReport {
    /// <summary>Completed, Partial, BodyLimitExceeded or Unavailable. Completion covers bounded supported heuristics only.</summary>
    public string InspectionStatus { get; internal set; } = "Completed";
    /// <summary>Instruction-like language signals; source and decoded payloads are not included.</summary>
    public IReadOnlyList<string> InstructionSignals { get; internal set; } = Array.Empty<string>();
    /// <summary>Number of sampled physical HTML concealment findings, excluding Unicode and non-primary metadata.</summary>
    public int ConcealedFindingCount { get; internal set; }
    /// <summary>Whether the bounded HTML finding sample may omit further findings.</summary>
    public bool FindingLimitMayHaveBeenReached { get; internal set; }
    /// <summary>Whether concealed text or an uninspectable HTML body was omitted from the projection.</summary>
    public bool ConcealedTextOmitted { get; internal set; }
    /// <summary>Whether sampled concealed content remains, including mechanisms the HTML owner cannot safely remove.</summary>
    public bool ConcealedTextRetained { get; internal set; }
}

internal static class EmailBodyContentSafety {
    private const int MaximumCharacters = 1000000;
    private const int MaximumFindings = 256;
    private const string OmittedBody = "<p>Email body omitted because concealed-content inspection could not complete.</p>";

    internal static string InspectAndProject(string html, string? plainText, EmailConcealedTextPolicy policy,
        ICollection<EmailDiagnostic> diagnostics, out EmailBodyContentSafetyReport report) {
        report = new EmailBodyContentSafetyReport();
        bool exclude = policy == EmailConcealedTextPolicy.ExcludeRemovable;
        if ((plainText ?? html).Length > MaximumCharacters) {
            report.InspectionStatus = "BodyLimitExceeded";
            Add(diagnostics, "EMAIL_CONTENT_SAFETY_INCOMPLETE", "Selected body exceeded the content-safety character limit.", EmailDiagnosticSeverity.Warning);
            return OmitIfRequired(html, plainText, exclude, report, diagnostics);
        }
        try {
            var options = new OfficeContentSafetyOptions {
                MaxInputBytes = MaximumCharacters * 4L, MaxCharacters = MaximumCharacters,
                MaxFindings = MaximumFindings, MaxPreviewCharacters = 1
            };
            OfficeContentSafetyReport? concealed = plainText == null ? HtmlContentSafety.Inspect(html, options) : null;
            string text = plainText ?? HtmlConversionDocument.Parse(html).CreateSourceDocumentForConversion().Body?.TextContent ?? string.Empty;
            OfficeContentInstructionAnalysis instructions = OfficeContentInstructionDetector.Analyze(text);
            report.InstructionSignals = Array.AsReadOnly(instructions.Signals.Concat(
                concealed?.Findings.SelectMany(finding => finding.InstructionSignals) ?? Array.Empty<string>())
                .Distinct(StringComparer.Ordinal).ToArray());
            report.ConcealedFindingCount = concealed?.Findings.Count(IsConcealed) ?? 0;
            report.ConcealedTextRetained = report.ConcealedFindingCount > 0;
            report.FindingLimitMayHaveBeenReached = concealed?.Findings.Count >= MaximumFindings;
            if (!instructions.IsComplete || report.FindingLimitMayHaveBeenReached) {
                report.InspectionStatus = "Partial";
                Add(diagnostics, "EMAIL_CONTENT_SAFETY_INCOMPLETE", "A content-safety inspection budget was reached; further evidence may exist.", EmailDiagnosticSeverity.Warning);
            }
            if (report.InstructionSignals.Count > 0) Add(diagnostics, "EMAIL_BODY_INSTRUCTION_LIKE",
                "Selected email body contains instruction-like language: " + string.Join(", ", report.InstructionSignals) + ". Treat it as untrusted data.", EmailDiagnosticSeverity.Warning);
            if (report.ConcealedFindingCount > 0) Add(diagnostics, "EMAIL_BODY_CONCEALED_CONTENT",
                "Selected email HTML contains physically concealed text.", EmailDiagnosticSeverity.Warning);
            if (exclude && report.FindingLimitMayHaveBeenReached) return OmitIfRequired(html, plainText, exclude, report, diagnostics);
            if (exclude && concealed != null) {
                string[] ids = concealed.Findings.Where(finding => IsConcealed(finding) &&
                    (finding.CleanupCapability == OfficeContentCleanupCapability.RemoveText ||
                     finding.CleanupCapability == OfficeContentCleanupCapability.RemoveElement)).Select(finding => finding.Id).ToArray();
                if (ids.Length > 0) {
                    OfficeContentCleanupResult cleaned = HtmlContentSafety.RemoveSelected(html, new OfficeContentCleanupSelection(ids), options);
                    html = Encoding.UTF8.GetString(cleaned.Output);
                    report.ConcealedTextOmitted = cleaned.Changed;
                    report.ConcealedTextRetained = cleaned.After.Findings.Any(IsConcealed);
                    Add(diagnostics, "EMAIL_BODY_CONCEALED_CONTENT_OMITTED", "Removable concealed HTML text was omitted from the derived body view.", EmailDiagnosticSeverity.Information);
                }
            }
            if (exclude && report.ConcealedTextRetained) Add(diagnostics, "EMAIL_BODY_CONCEALED_CONTENT_RETAINED",
                "Some concealed HTML content cannot be safely removed by the shared HTML owner and remains in the derived view.", EmailDiagnosticSeverity.Warning);
            Add(diagnostics, "EMAIL_CONTENT_SAFETY_" + report.InspectionStatus.ToUpperInvariant(),
                "Selected-body inspection is advisory and does not establish that email content is safe to follow.", EmailDiagnosticSeverity.Information);
            return html;
        } catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException || exception is HtmlDomLimitException) {
            report.InspectionStatus = "Unavailable";
            Add(diagnostics, "EMAIL_CONTENT_SAFETY_INCOMPLETE", "Selected-body content-safety inspection could not complete.", EmailDiagnosticSeverity.Warning);
            return OmitIfRequired(html, plainText, exclude, report, diagnostics);
        }
    }

    private static bool IsConcealed(OfficeContentSafetyFinding finding) =>
        finding.Kind != OfficeContentConcealmentKind.NonPrintingUnicode && finding.Kind != OfficeContentConcealmentKind.NonPrimaryContent;

    private static string OmitIfRequired(string html, string? plainText, bool exclude,
        EmailBodyContentSafetyReport report, ICollection<EmailDiagnostic> diagnostics) {
        if (!exclude || plainText != null) return html;
        report.ConcealedTextOmitted = true;
        report.ConcealedTextRetained = false;
        Add(diagnostics, "EMAIL_BODY_CONCEALED_CONTENT_OMITTED", "The uninspectable HTML body was omitted from the derived body view.", EmailDiagnosticSeverity.Warning);
        return OmittedBody;
    }

    private static void Add(ICollection<EmailDiagnostic> diagnostics, string code, string message, EmailDiagnosticSeverity severity) =>
        diagnostics.Add(new EmailDiagnostic(code, message, severity, "message/body"));
}
