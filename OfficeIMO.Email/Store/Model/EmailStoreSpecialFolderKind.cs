namespace OfficeIMO.Email.Store;

/// <summary>Well-known role assigned to a store folder when the source identifies one.</summary>
public enum EmailStoreSpecialFolderKind {
    /// <summary>No well-known role could be established.</summary>
    Unknown = 0,
    /// <summary>Message-store root.</summary>
    Root = 1,
    /// <summary>Root of the interpersonal-message hierarchy.</summary>
    IpmSubtree = 2,
    /// <summary>Default received-message folder.</summary>
    Inbox = 3,
    /// <summary>Outgoing-message queue.</summary>
    Outbox = 4,
    /// <summary>Default sent-message folder.</summary>
    SentItems = 5,
    /// <summary>Default deleted-item folder.</summary>
    DeletedItems = 6,
    /// <summary>Default draft-message folder.</summary>
    Drafts = 7,
    /// <summary>Default calendar folder.</summary>
    Calendar = 8,
    /// <summary>Default contacts folder.</summary>
    Contacts = 9,
    /// <summary>Default tasks folder.</summary>
    Tasks = 10,
    /// <summary>Default notes folder.</summary>
    Notes = 11,
    /// <summary>Default journal folder.</summary>
    Journal = 12,
    /// <summary>Junk or spam folder.</summary>
    JunkEmail = 13,
    /// <summary>Search-folder root.</summary>
    SearchRoot = 14,
    /// <summary>Shared view definitions.</summary>
    CommonViews = 15,
    /// <summary>Personal view definitions.</summary>
    PersonalViews = 16,
    /// <summary>Archive folder.</summary>
    Archive = 17,
    /// <summary>Synchronization-issue parent folder.</summary>
    SyncIssues = 18,
    /// <summary>Synchronization conflicts.</summary>
    Conflicts = 19,
    /// <summary>Local synchronization failures.</summary>
    LocalFailures = 20,
    /// <summary>Server synchronization failures.</summary>
    ServerFailures = 21,
    /// <summary>RSS feed folder.</summary>
    RssFeeds = 22,
    /// <summary>Reminder search folder.</summary>
    Reminders = 23,
    /// <summary>To-do search folder.</summary>
    ToDo = 24
}
