namespace OfficeIMO.Email.Store;

/// <summary>Aggregate, privacy-safe proof produced by reopening and comparing a converted PST.</summary>
public sealed class EmailStorePstVerificationReport {
    internal EmailStorePstVerificationReport(int attemptedItems, int matchedItems,
        int mismatchedItems, int failedItems, IReadOnlyList<EmailStorePstVerificationIssue> issues,
        bool issuesTruncated, string? manifestPath) {
        AttemptedItems = attemptedItems;
        MatchedItems = matchedItems;
        MismatchedItems = mismatchedItems;
        FailedItems = failedItems;
        Issues = issues;
        IssuesTruncated = issuesTruncated;
        ManifestPath = manifestPath;
    }

    /// <summary>Number of written items inspected.</summary>
    public int AttemptedItems { get; }

    /// <summary>Number of destination items matching the source semantic projection.</summary>
    public int MatchedItems { get; }

    /// <summary>Number of readable destination items with semantic differences.</summary>
    public int MismatchedItems { get; }

    /// <summary>Number of items that could not be read or compared.</summary>
    public int FailedItems { get; }

    /// <summary>Bounded mismatch/failure details. Matching items are summarized rather than retained.</summary>
    public IReadOnlyList<EmailStorePstVerificationIssue> Issues { get; }

    /// <summary>True when additional issues existed beyond the configured report limit.</summary>
    public bool IssuesTruncated { get; }

    /// <summary>Committed aggregate manifest path, or null when no manifest was requested.</summary>
    public string? ManifestPath { get; }

    /// <summary>True when every written item was reopened and matched.</summary>
    public bool IsSuccessful => AttemptedItems == MatchedItems && MismatchedItems == 0 && FailedItems == 0;

    internal EmailStorePstVerificationReport WithManifestPath(string manifestPath) =>
        new EmailStorePstVerificationReport(AttemptedItems, MatchedItems, MismatchedItems,
            FailedItems, Issues, IssuesTruncated, manifestPath);
}
