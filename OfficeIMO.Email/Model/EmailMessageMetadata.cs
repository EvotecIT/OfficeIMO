namespace OfficeIMO.Email;

/// <summary>Common Outlook/MAPI message metadata not represented by MIME headers alone.</summary>
public sealed class EmailMessageMetadata {
    private readonly OutlookCategoryCollection _categories = new OutlookCategoryCollection();

    /// <summary>Outlook follow-up flag semantics.</summary>
    public OutlookFollowUp FollowUp { get; } = new OutlookFollowUp();

    /// <summary>Outlook reminder semantics for message and contact items.</summary>
    public OutlookReminder Reminder { get; } = new OutlookReminder();

    /// <summary>Outlook voting options and response semantics.</summary>
    public OutlookVoting Voting { get; } = new OutlookVoting();

    /// <summary>Subject prefix such as <c>RE: </c> or <c>FW: </c>.</summary>
    public string? SubjectPrefix { get; set; }

    /// <summary>Subject without the reply/forward prefix.</summary>
    public string? NormalizedSubject { get; set; }

    /// <summary>Conversation topic displayed by Outlook.</summary>
    public string? ConversationTopic { get; set; }

    /// <summary>Binary Outlook conversation index.</summary>
    public byte[]? ConversationIndex { get; set; }

    /// <summary>Internet References field retained in MAPI form.</summary>
    public string? InternetReferences { get; set; }

    /// <summary>Internet In-Reply-To identifier retained in MAPI form.</summary>
    public string? InReplyToId { get; set; }

    /// <summary>Message importance.</summary>
    public EmailMessageImportance? Importance { get; set; }

    /// <summary>Message transport priority.</summary>
    public EmailMessagePriority? Priority { get; set; }

    /// <summary>Outlook icon-index hint.</summary>
    public int? IconIndex { get; set; }

    /// <summary>Whether the message is still being composed.</summary>
    public bool IsDraft { get; set; }

    /// <summary>Read state when present.</summary>
    public bool? IsRead { get; set; }

    /// <summary>Whether a read receipt was requested.</summary>
    public bool ReadReceiptRequested { get; set; }

    /// <summary>MIME destination for a requested read receipt.</summary>
    public string? ReadReceiptDestination { get; set; }

    /// <summary>Whether a delivery receipt was requested.</summary>
    public bool DeliveryReceiptRequested { get; set; }

    /// <summary>MIME destination for a requested delivery receipt.</summary>
    public string? DeliveryReceiptDestination { get; set; }

    /// <summary>MAPI sensitivity value.</summary>
    public int? Sensitivity { get; set; }

    /// <summary>Original MAPI sensitivity value before a change.</summary>
    public int? OriginalSensitivity { get; set; }

    /// <summary>Name stamped as the last modifier.</summary>
    public string? LastModifierName { get; set; }

    /// <summary>Windows locale identifier stamped on the message.</summary>
    public int? LocaleId { get; set; }

    /// <summary>Declared source message size.</summary>
    public int? DeclaredSize { get; set; }

    /// <summary>Binary conversation identifier.</summary>
    public byte[]? ConversationId { get; set; }

    /// <summary>Outlook editor-format numeric value.</summary>
    public int? EditorFormat { get; set; }

    /// <summary>Raw current-reactions summary blob retained for Outlook-compatible consumers.</summary>
    public byte[]? ReactionsSummary { get; set; }

    /// <summary>Raw owner-reaction history blob retained for Outlook-compatible consumers.</summary>
    public byte[]? OwnerReactionHistory { get; set; }

    /// <summary>Owner's current reaction type when present.</summary>
    public string? OwnerReactionType { get; set; }

    /// <summary>Time of the owner's current reaction.</summary>
    public DateTimeOffset? OwnerReactionTime { get; set; }

    /// <summary>Declared current reaction count.</summary>
    public int? ReactionsCount { get; set; }

    /// <summary>Message creation timestamp.</summary>
    public DateTimeOffset? CreatedDate { get; set; }

    /// <summary>Last MAPI modification timestamp.</summary>
    public DateTimeOffset? ModifiedDate { get; set; }

    /// <summary>Outlook categories.</summary>
    public OutlookCategoryCollection Categories => _categories;
}
