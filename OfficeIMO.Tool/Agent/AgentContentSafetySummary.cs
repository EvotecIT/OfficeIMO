using OfficeIMO.Email;
using System.Text.Json.Serialization;

namespace OfficeIMO.Tool.Agent;

/// <summary>Stable advisory evidence retained even when diagnostic samples are removed for the output budget.</summary>
public sealed class AgentContentSafetySummary {
    /// <summary>NotInspected, Completed or Partial. These describe bounded email-body inspection, never a safety verdict.</summary>
    public string Status { get; set; } = "NotInspected";
    /// <summary>Whether instruction-like language was observed.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool InstructionLike { get; set; }
    /// <summary>Retained, Omitted or PartlyOmitted across the inspected bodies; absent when no concealment was observed.</summary>
    public string? ConcealedText { get; set; }

    internal void Include(EmailBodyContentSafetyReport report) {
        AgentContentSafetySummary next = FromReport(report);
        Status = Status == "Partial" || next.Status == "Partial" ? "Partial" : "Completed";
        InstructionLike |= next.InstructionLike;
        if (next.ConcealedText != null) ConcealedText = ConcealedText == null || ConcealedText == next.ConcealedText
            ? next.ConcealedText : "PartlyOmitted";
    }

    internal static AgentContentSafetySummary FromDiagnostics(IEnumerable<string> codes) {
        var values = new HashSet<string>(codes, StringComparer.Ordinal);
        bool omitted = values.Contains("EMAIL_BODY_CONCEALED_CONTENT_OMITTED");
        bool retained = values.Contains("EMAIL_BODY_CONCEALED_CONTENT_RETAINED") ||
            values.Contains("EMAIL_BODY_CONCEALED_CONTENT") && !omitted;
        return new AgentContentSafetySummary {
            Status = values.Contains("EMAIL_CONTENT_SAFETY_INCOMPLETE") ? "Partial" :
                values.Contains("EMAIL_CONTENT_SAFETY_COMPLETED") ? "Completed" : "NotInspected",
            InstructionLike = values.Contains("EMAIL_BODY_INSTRUCTION_LIKE"),
            ConcealedText = omitted ? retained ? "PartlyOmitted" : "Omitted" : retained ? "Retained" : null
        };
    }

    internal static AgentContentSafetySummary FromReport(EmailBodyContentSafetyReport? report) => report == null
        ? new AgentContentSafetySummary()
        : new AgentContentSafetySummary {
            Status = report.InspectionStatus == "Completed" ? "Completed" : "Partial",
            InstructionLike = report.InstructionSignals.Count > 0,
            ConcealedText = report.ConcealedTextOmitted ? report.ConcealedTextRetained ? "PartlyOmitted" : "Omitted" :
                report.ConcealedTextRetained ? "Retained" : null
        };
}
