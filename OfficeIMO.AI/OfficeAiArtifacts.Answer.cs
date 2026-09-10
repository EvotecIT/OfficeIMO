using System.Text;

namespace OfficeIMO.AI;

public static partial class OfficeAiArtifacts {
    /// <summary>
    /// Formats an Ask, Explain or Summarize result as plain text with source identity, quotes and limitations.
    /// The question and display name are caller-owned labels; no source file is opened and no output is published.
    /// </summary>
    public static string FormatAnswer(OfficeAiDocument document, OfficeAiResult result, string question, string sourceName) {
        CheckSource(document, result);
        ArgumentNullException.ThrowIfNull(question);
        ArgumentNullException.ThrowIfNull(sourceName);
        if (result.Operation is not (OfficeAiOperation.Ask or OfficeAiOperation.Explain or OfficeAiOperation.Summarize))
            throw new ArgumentException("This artifact represents answers and summaries, not structured extraction or parsing.", nameof(result));
        var text = new StringBuilder();
        text.AppendLine("Document answer — requires review");
        text.Append("Document: ").AppendLine(sourceName);
        text.Append("Source SHA-256: ").AppendLine(document.SourceHash);
        text.Append("Evidence snapshot SHA-256: ").AppendLine(document.SnapshotHash);
        text.Append("Provider: ").Append(result.Profile.Provider).Append("; model: ").AppendLine(result.Profile.Model);
        text.Append("Operation: ").Append(result.Operation).Append("; outcome: ").AppendLine(result.Status.ToString());
        text.Append("Request: ").AppendLine(result.RequestId);
        text.AppendLine().AppendLine("Question:").AppendLine(question);
        foreach (OfficeAiClaim claim in result.Claims) {
            text.AppendLine().AppendLine("Answer:").AppendLine(claim.Text);
            foreach (OfficeAiCitation citation in claim.Citations) {
                text.Append("Source: ").Append(citation.EvidenceId);
                if (citation.Page.HasValue) text.Append("; page ").Append(citation.Page.Value);
                text.AppendLine();
                if (!string.IsNullOrWhiteSpace(citation.Quote)) text.Append("Quote: ").AppendLine(citation.Quote);
            }
        }
        if (result.Claims.Count == 0) text.AppendLine().AppendLine("No supported answer was returned.");
        text.AppendLine().AppendLine("Limitations:");
        text.AppendLine("A matching source quote does not prove the interpretation. Review the answer against the document.");
        text.Append("Omitted evidence items: ").AppendLine(result.OmittedEvidenceIds.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (result.EmptyPages.Count > 0) text.Append("Pages without evidence: ").AppendLine(string.Join(", ", result.EmptyPages));
        if (document.HasSourceDiagnostics) text.AppendLine("The source reader reported limitations; complete reconstruction is not certified.");
        if (result.Diagnostics.Count > 0) text.Append("Diagnostics: ").AppendLine(string.Join(", ", result.Diagnostics));
        return text.ToString();
    }
}
