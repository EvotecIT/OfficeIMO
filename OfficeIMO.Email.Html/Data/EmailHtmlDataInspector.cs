using OfficeIMO.ContentSafety;
using OfficeIMO.Email.Data;
using OfficeIMO.Html;

namespace OfficeIMO.Email;

/// <summary>Bounded metadata and HTML-policy evidence without body previews or attachment payloads.</summary>
public sealed class EmailHtmlDataInspectionReport {
    internal EmailHtmlDataInspectionReport(EmailDataInspectionReport metadata) { Metadata = metadata; }
    /// <summary>Snapshot from the canonical email-data owner.</summary>
    public EmailDataInspectionReport Metadata { get; }
    /// <summary>Completed, BodyLimitExceeded, InspectionUnavailable, NoHtmlBody or CatalogOnly. Only Completed qualifies the HTML sample.</summary>
    public string HtmlInspectionStatus { get; internal set; } = "CatalogOnly";
    /// <summary>Generic diagnostic code when the HTML engine cannot complete inspection; raw exception text is omitted.</summary>
    public string? HtmlDiagnosticCode { get; internal set; }
    /// <summary>Original HTML elements removed by the shared body projection's active-content policy.</summary>
    public int BlockedElementCount { get; internal set; }
    /// <summary>Original event-handler attributes removed by that policy.</summary>
    public int EventHandlerAttributeCount { get; internal set; }
    /// <summary>Bounded concealed-text findings without original text, content hashes or previews.</summary>
    public IReadOnlyList<EmailHtmlInspectionFinding> Findings { get; internal set; } = Array.Empty<EmailHtmlInspectionFinding>();
    /// <summary>Whether the finding sample reached its bound; additional findings may exist.</summary>
    public bool FindingLimitMayHaveBeenReached { get; internal set; }
    /// <summary>Advisory selected-body evidence, including plain text and bounded inline Base64. Null for catalog-only artifacts.</summary>
    public EmailBodyContentSafetyReport? BodyContentSafety { get; internal set; }
}

/// <summary>One concealed-text mechanism and risk classification, without the private text.</summary>
public sealed class EmailHtmlInspectionFinding {
    internal EmailHtmlInspectionFinding(string kind, string risk, bool instructionLike) { Kind = kind; Risk = risk; IsInstructionLike = instructionLike; }
    /// <summary>Shared HTML concealment mechanism.</summary>
    public string Kind { get; }
    /// <summary>Shared content-safety risk classification; this is not a malware verdict.</summary>
    public string Risk { get; }
    /// <summary>Whether bounded heuristics found instruction-like concealed text.</summary>
    public bool IsInstructionLike { get; }
}

/// <summary>Optional HTML adapter over EmailDataArtifact and the existing HTML/content-safety engines.</summary>
public static class EmailHtmlDataInspector {
    /// <summary>Opens once, inspects bounded metadata and original HTML, then closes resources. No network or mutation occurs.</summary>
    public static EmailHtmlDataInspectionReport Inspect(string path, EmailDataInspectionOptions? options = null,
        int maxHtmlCharacters = 1000000, CancellationToken cancellationToken = default) {
        ValidateLimit(maxHtmlCharacters);
        var policy = options ?? new EmailDataInspectionOptions();
        using var opened = EmailDataArtifact.Open(path, policy.OpenOptions, cancellationToken);
        return Inspect(opened, policy, maxHtmlCharacters, cancellationToken);
    }

    /// <summary>Inspects an existing artifact without disposing it. The original body and attachment streams are left intact.</summary>
    public static EmailHtmlDataInspectionReport Inspect(EmailDataOpenResult opened, EmailDataInspectionOptions? options = null,
        int maxHtmlCharacters = 1000000, CancellationToken cancellationToken = default) {
        ValidateLimit(maxHtmlCharacters);
        var policy = options ?? new EmailDataInspectionOptions();
        var report = new EmailHtmlDataInspectionReport(EmailDataInspector.Inspect(opened, policy, cancellationToken));
        if (opened.EmailDocument == null) return report;
        cancellationToken.ThrowIfCancellationRequested();
        string selected = !string.IsNullOrWhiteSpace(opened.EmailDocument.Body.Html) ? opened.EmailDocument.Body.Html! :
            !string.IsNullOrWhiteSpace(opened.EmailDocument.Body.Rtf) ? opened.EmailDocument.Body.Rtf! : opened.EmailDocument.Body.Text ?? string.Empty;
        if (selected.Length > Math.Min(maxHtmlCharacters, 1000000)) {
            report.BodyContentSafety = new EmailBodyContentSafetyReport { InspectionStatus = "BodyLimitExceeded" };
        } else {
            try {
                report.BodyContentSafety = EmailBodyProjection.Create(opened.EmailDocument, new EmailBodyProjectionOptions {
                    IncludeResources = false, IncludeResourceReferences = false, InspectContentSafety = true
                }).ContentSafety;
            } catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException || exception is HtmlDomLimitException) {
                report.BodyContentSafety = new EmailBodyContentSafetyReport { InspectionStatus = "Unavailable" };
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        string? html = opened.EmailDocument.Body.Html;
        if (html == null) { report.HtmlInspectionStatus = "NoHtmlBody"; return report; }
        if (html.Length > maxHtmlCharacters) { report.HtmlInspectionStatus = "BodyLimitExceeded"; return report; }
        cancellationToken.ThrowIfCancellationRequested();
        try {
            OfficeContentSafetyReport safety = HtmlContentSafety.Inspect(html, new OfficeContentSafetyOptions {
                MaxInputBytes = checked((long)maxHtmlCharacters * 4), MaxCharacters = maxHtmlCharacters,
                MaxFindings = policy.MaxSamples, MaxPreviewCharacters = 1
            });
            cancellationToken.ThrowIfCancellationRequested();
            var native = HtmlConversionDocument.Parse(html).CreateSourceDocumentForConversion();
            var activity = EmailHtmlActiveContent.Inspect(native);
            cancellationToken.ThrowIfCancellationRequested();
            report.BlockedElementCount = activity.Elements; report.EventHandlerAttributeCount = activity.Attributes;
            report.Findings = Array.AsReadOnly(safety.Findings.Select(finding => new EmailHtmlInspectionFinding(
                finding.Kind.ToString(), finding.Risk.ToString(), finding.IsInstructionLike)).ToArray());
            report.FindingLimitMayHaveBeenReached = safety.Findings.Count >= policy.MaxSamples;
            report.HtmlInspectionStatus = "Completed";
        } catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException || exception is HtmlDomLimitException) {
            cancellationToken.ThrowIfCancellationRequested();
            report.HtmlInspectionStatus = "InspectionUnavailable";
            report.HtmlDiagnosticCode = "EMAIL_HTML_INSPECTION_UNAVAILABLE";
        }
        return report;
    }
    private static void ValidateLimit(int maxHtmlCharacters) {
        if (maxHtmlCharacters <= 0 || maxHtmlCharacters > 2000000) throw new ArgumentOutOfRangeException(nameof(maxHtmlCharacters));
    }
}
