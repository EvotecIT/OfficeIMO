using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>
/// The finished states every Studio queue presents the same way, so a colour, icon or word means the same thing in
/// Convert, OCR, Jobs and the completion toast. Output that may exist but could not be confirmed needs checking; it is
/// neither a success nor a failure.
/// </summary>
public enum StudioOutcome {
    /// <summary>Waiting or still running.</summary>
    None,
    /// <summary>Completed and verified with nothing to report.</summary>
    Succeeded,
    /// <summary>Completed, with warnings or approximated or omitted content to review.</summary>
    CompletedWithNotes,
    /// <summary>Stopped by the user; outputs that were already published remain.</summary>
    Cancelled,
    /// <summary>The runner stopped without confirming the output; inspect the destination before retrying.</summary>
    NeedsChecking,
    /// <summary>Did not produce the requested output.</summary>
    Failed
}

internal static class StudioOutcomes {
    internal static StudioOutcome From(OfficeWorkflowStatus? status, bool hasNotes = false) => status switch {
        null => StudioOutcome.None,
        OfficeWorkflowStatus.Completed => hasNotes ? StudioOutcome.CompletedWithNotes : StudioOutcome.Succeeded,
        OfficeWorkflowStatus.Cancelled => StudioOutcome.Cancelled,
        OfficeWorkflowStatus.Unconfirmed => StudioOutcome.NeedsChecking,
        _ => StudioOutcome.Failed
    };

    /// <summary>The theme brush family for an outcome: Success, Warning, Error, or the muted text tone.</summary>
    internal static string Tone(StudioOutcome outcome) => outcome switch {
        StudioOutcome.Succeeded => "Success",
        StudioOutcome.CompletedWithNotes or StudioOutcome.NeedsChecking => "Warning",
        StudioOutcome.Failed => "Error",
        _ => "Accent"
    };
}
