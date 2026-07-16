namespace OfficeIMO.Email.Store;

internal static class EmailStoreSpecialFolderClassifier {
    internal static EmailStoreSpecialFolderKind FromDisplayName(string? name) {
        if (string.IsNullOrWhiteSpace(name)) return EmailStoreSpecialFolderKind.Unknown;
        string value = name!.Trim();
        if (EqualsAny(value, "Inbox")) return EmailStoreSpecialFolderKind.Inbox;
        if (EqualsAny(value, "Outbox")) return EmailStoreSpecialFolderKind.Outbox;
        if (EqualsAny(value, "Sent", "Sent Items", "Sent Messages")) return EmailStoreSpecialFolderKind.SentItems;
        if (EqualsAny(value, "Deleted Items", "Trash", "Deleted Messages")) return EmailStoreSpecialFolderKind.DeletedItems;
        if (EqualsAny(value, "Drafts", "Draft")) return EmailStoreSpecialFolderKind.Drafts;
        if (EqualsAny(value, "Calendar")) return EmailStoreSpecialFolderKind.Calendar;
        if (EqualsAny(value, "Contacts")) return EmailStoreSpecialFolderKind.Contacts;
        if (EqualsAny(value, "Tasks")) return EmailStoreSpecialFolderKind.Tasks;
        if (EqualsAny(value, "Notes")) return EmailStoreSpecialFolderKind.Notes;
        if (EqualsAny(value, "Journal")) return EmailStoreSpecialFolderKind.Journal;
        if (EqualsAny(value, "Junk", "Junk E-mail", "Spam")) return EmailStoreSpecialFolderKind.JunkEmail;
        if (EqualsAny(value, "Archive")) return EmailStoreSpecialFolderKind.Archive;
        if (EqualsAny(value, "Sync Issues")) return EmailStoreSpecialFolderKind.SyncIssues;
        if (EqualsAny(value, "Conflicts")) return EmailStoreSpecialFolderKind.Conflicts;
        if (EqualsAny(value, "Local Failures")) return EmailStoreSpecialFolderKind.LocalFailures;
        if (EqualsAny(value, "Server Failures")) return EmailStoreSpecialFolderKind.ServerFailures;
        if (EqualsAny(value, "RSS Feeds")) return EmailStoreSpecialFolderKind.RssFeeds;
        if (EqualsAny(value, "Reminders")) return EmailStoreSpecialFolderKind.Reminders;
        if (EqualsAny(value, "To-Do", "To Do")) return EmailStoreSpecialFolderKind.ToDo;
        return EmailStoreSpecialFolderKind.Unknown;
    }

    private static bool EqualsAny(string value, params string[] candidates) =>
        candidates.Any(candidate => string.Equals(value, candidate, StringComparison.OrdinalIgnoreCase));
}
