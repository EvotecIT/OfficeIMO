namespace OfficeIMO.Pdf;

/// <summary>Standard review state stored on a PDF annotation.</summary>
public enum PdfAnnotationReviewState {
    /// <summary>No decision has been recorded.</summary>
    None,
    /// <summary>The annotation has been accepted.</summary>
    Accepted,
    /// <summary>The annotation has been rejected.</summary>
    Rejected,
    /// <summary>The review item has been cancelled.</summary>
    Cancelled,
    /// <summary>The review item has been completed.</summary>
    Completed,
    /// <summary>The annotation is marked.</summary>
    Marked,
    /// <summary>The annotation is not marked.</summary>
    Unmarked
}

/// <summary>Reply, state, subject, and intent metadata read from a PDF annotation.</summary>
public sealed class PdfAnnotationReviewInfo {
    internal PdfAnnotationReviewInfo(
        int? inReplyToObjectNumber,
        string? replyType,
        string? state,
        string? stateModel,
        string? subject,
        string? intent) {
        InReplyToObjectNumber = inReplyToObjectNumber;
        ReplyType = replyType;
        State = state;
        StateModel = stateModel;
        Subject = subject;
        Intent = intent;
    }

    /// <summary>Indirect object number of the annotation referenced by /IRT.</summary>
    public int? InReplyToObjectNumber { get; }

    /// <summary>Reply relationship from /RT, normally R or Group.</summary>
    public string? ReplyType { get; }

    /// <summary>Raw annotation state name from /State.</summary>
    public string? State { get; }

    /// <summary>Raw state model name from /StateModel.</summary>
    public string? StateModel { get; }

    /// <summary>Annotation subject from /Subj.</summary>
    public string? Subject { get; }

    /// <summary>Annotation intent name from /IT.</summary>
    public string? Intent { get; }

    /// <summary>True when this annotation has an absent/default or explicit R reply relationship.</summary>
    public bool IsReply => InReplyToObjectNumber.HasValue &&
        (string.IsNullOrEmpty(ReplyType) || string.Equals(ReplyType, "R", StringComparison.Ordinal));

    /// <summary>True when this annotation declares a Group relationship rather than a conversational reply.</summary>
    public bool IsGroup => InReplyToObjectNumber.HasValue && string.Equals(ReplyType, "Group", StringComparison.Ordinal);

    /// <summary>Typed standard state when the raw state and model form a known combination.</summary>
    public PdfAnnotationReviewState? StandardState => (StateModel, State) switch {
        ("Review", "None") => PdfAnnotationReviewState.None,
        ("Review", "Accepted") => PdfAnnotationReviewState.Accepted,
        ("Review", "Rejected") => PdfAnnotationReviewState.Rejected,
        ("Review", "Cancelled") => PdfAnnotationReviewState.Cancelled,
        ("Review", "Completed") => PdfAnnotationReviewState.Completed,
        ("Marked", "Marked") => PdfAnnotationReviewState.Marked,
        ("Marked", "Unmarked") => PdfAnnotationReviewState.Unmarked,
        _ => null
    };
}
