using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Kind of logical entity checked after a staged PST mutation.</summary>
public enum EmailStorePstMutationVerificationEntity {
    /// <summary>A folder and its parent/name/container-class projection.</summary>
    Folder = 0,
    /// <summary>An email or typed Outlook item.</summary>
    Item = 1
}

/// <summary>Privacy-safe detail for one folder or item that did not pass mutation verification.</summary>
public sealed class EmailStorePstMutationVerificationIssue {
    internal EmailStorePstMutationVerificationIssue(
        EmailStorePstMutationVerificationEntity entity, string entityId,
        string destinationEntityId, string code,
        IReadOnlyList<EmailSemanticDifference> differences) {
        Entity = entity;
        EntityId = entityId;
        DestinationEntityId = destinationEntityId;
        Code = code;
        Differences = differences;
    }

    /// <summary>Kind of entity that was checked.</summary>
    public EmailStorePstMutationVerificationEntity Entity { get; }

    /// <summary>Source or transaction-local identifier.</summary>
    public string EntityId { get; }

    /// <summary>Identifier assigned in the rewritten PST.</summary>
    public string DestinationEntityId { get; }

    /// <summary>Stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Value-free semantic difference paths.</summary>
    public IReadOnlyList<EmailSemanticDifference> Differences { get; }
}

/// <summary>Aggregate semantic verification for a staged PST mutation.</summary>
public sealed class EmailStorePstMutationVerificationReport {
    internal EmailStorePstMutationVerificationReport(
        int attemptedFolders, int matchedFolders, int mismatchedFolders, int failedFolders,
        int attemptedItems, int matchedItems,
        int mismatchedItems, int failedItems,
        IReadOnlyList<EmailStorePstMutationVerificationIssue> issues, bool issuesTruncated) {
        AttemptedFolders = attemptedFolders;
        MatchedFolders = matchedFolders;
        MismatchedFolders = mismatchedFolders;
        FailedFolders = failedFolders;
        AttemptedItems = attemptedItems;
        MatchedItems = matchedItems;
        MismatchedItems = mismatchedItems;
        FailedItems = failedItems;
        Issues = issues;
        IssuesTruncated = issuesTruncated;
    }

    /// <summary>Number of intended folders considered.</summary>
    public int AttemptedFolders { get; }

    /// <summary>Number of intended folders with the expected hierarchy and metadata.</summary>
    public int MatchedFolders { get; }

    /// <summary>Number of readable resulting folders with structural differences.</summary>
    public int MismatchedFolders { get; }

    /// <summary>Number of intended folders that could not be found or checked.</summary>
    public int FailedFolders { get; }

    /// <summary>Number of resulting items considered.</summary>
    public int AttemptedItems { get; }

    /// <summary>Number of resulting items matching the intended semantic projection.</summary>
    public int MatchedItems { get; }

    /// <summary>Number of readable resulting items with semantic differences.</summary>
    public int MismatchedItems { get; }

    /// <summary>Number of items that could not be reopened or compared.</summary>
    public int FailedItems { get; }

    /// <summary>Bounded value-free mismatch and failure details.</summary>
    public IReadOnlyList<EmailStorePstMutationVerificationIssue> Issues { get; }

    /// <summary>Whether additional issue details were omitted.</summary>
    public bool IssuesTruncated { get; }

    /// <summary>True when every resulting item was reopened and matched.</summary>
    public bool IsSuccessful => MismatchedFolders == 0 && FailedFolders == 0 &&
        MismatchedItems == 0 && FailedItems == 0;
}
