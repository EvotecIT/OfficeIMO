namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Builds a bounded cross-folder conversation graph from Internet threading fields, Outlook conversation
    /// identities, and meeting/task lifecycle identifiers. Subject-only links remain explicitly heuristic.
    /// </summary>
    public EmailConversationGraph BuildConversationGraph(
        EmailConversationGraphOptions? options = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        return new EmailConversationGraphBuilder(this,
            options ?? new EmailConversationGraphOptions(), cancellationToken).Build();
    }
}
