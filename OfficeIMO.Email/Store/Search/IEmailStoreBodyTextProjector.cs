namespace OfficeIMO.Email.Store;

/// <summary>Optional format-owned body projection for semantic store search, without a dependency on an HTML engine.</summary>
public interface IEmailStoreBodyTextProjector {
    /// <summary>Stable identity of the projection policy and version. Checkpoints are bound to this identity.</summary>
    string Identity { get; }
    /// <summary>Projects one requested body field without modifying the message or opening attachments.
    /// The session enforces the returned-character bound; implementations must bound parsing and observe cancellation.</summary>
    string Project(EmailDocument document, EmailStoreContentSearchFields bodyField, int maxCharacters, CancellationToken cancellationToken);
}
