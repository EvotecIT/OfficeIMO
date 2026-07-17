using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Controls safe conversion of a supported mailbox store into a newly created Unicode PST.</summary>
public sealed class EmailStorePstConversionOptions {
    /// <summary>Creates conversion options.</summary>
    public EmailStorePstConversionOptions(
        bool overwriteExisting = false,
        bool failOnDataLoss = false,
        bool continueOnItemError = true,
        bool includeAssociatedItems = true,
        bool includeOrphanedItems = true,
        bool includeSearchFolders = true,
        int maxItems = int.MaxValue,
        int maxNestedMessageDepth = 32,
        string? displayName = null,
        bool verifyAfterWrite = true,
        EmailSemanticComparisonOptions? verificationOptions = null,
        string? verificationManifestPath = null,
        int maxVerificationIssues = 1_000) {
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        if (maxNestedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxNestedMessageDepth));
        if (maxVerificationIssues <= 0) throw new ArgumentOutOfRangeException(nameof(maxVerificationIssues));
        if (!string.IsNullOrWhiteSpace(verificationManifestPath) &&
            verificationOptions != null && !verificationOptions.UsesKeyedDigest) {
            throw new ArgumentException(
                "A persisted verification manifest requires keyed semantic fingerprints.",
                nameof(verificationOptions));
        }
        if (!verifyAfterWrite && !string.IsNullOrWhiteSpace(verificationManifestPath)) {
            throw new ArgumentException(
                "A verification manifest requires verifyAfterWrite to be enabled.",
                nameof(verificationManifestPath));
        }
        OverwriteExisting = overwriteExisting;
        FailOnDataLoss = failOnDataLoss;
        ContinueOnItemError = continueOnItemError;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        IncludeSearchFolders = includeSearchFolders;
        MaxItems = maxItems;
        MaxNestedMessageDepth = maxNestedMessageDepth;
        DisplayName = string.IsNullOrWhiteSpace(displayName) ? null : displayName;
        VerifyAfterWrite = verifyAfterWrite;
        VerificationOptions = verificationOptions;
        VerificationManifestPath = string.IsNullOrWhiteSpace(verificationManifestPath)
            ? null
            : verificationManifestPath;
        MaxVerificationIssues = maxVerificationIssues;
    }

    /// <summary>Whether an existing destination may be atomically replaced.</summary>
    public bool OverwriteExisting { get; }
    /// <summary>
    /// Whether any fidelity warning or error blocks completion. When verification is enabled, the destination is
    /// changed only after the staged PST passes semantic verification.
    /// </summary>
    public bool FailOnDataLoss { get; }
    /// <summary>Whether unreadable individual items are reported and skipped instead of aborting immediately.</summary>
    public bool ContinueOnItemError { get; }
    /// <summary>Whether folder-associated information items are copied.</summary>
    public bool IncludeAssociatedItems { get; }
    /// <summary>Whether items recovered from source indexes but absent from contents tables are copied.</summary>
    public bool IncludeOrphanedItems { get; }
    /// <summary>Whether search-folder results are copied as static folders and items.</summary>
    public bool IncludeSearchFolders { get; }
    /// <summary>Maximum source items inspected by one conversion.</summary>
    public int MaxItems { get; }
    /// <summary>Maximum embedded-message nesting depth written.</summary>
    public int MaxNestedMessageDepth { get; }
    /// <summary>Optional destination display name. The source display name is used when null.</summary>
    public string? DisplayName { get; }
    /// <summary>Whether the completed PST is reopened and compared item by item.</summary>
    public bool VerifyAfterWrite { get; }
    /// <summary>
    /// Optional semantic policy. A migration profile with an ephemeral HMAC key is used when null so report
    /// fingerprints cannot be correlated outside the conversion run.
    /// </summary>
    public EmailSemanticComparisonOptions? VerificationOptions { get; }
    /// <summary>
    /// Optional path for an aggregate TSV manifest containing ordinals, statuses, keyed source/destination
    /// fingerprints, an aggregate digest, and keyed difference-path tokens. Message subjects, addresses, content,
    /// filenames, and store identifiers are never written. Supply a keyed <see cref="VerificationOptions"/> policy
    /// to make the manifest reproducible; the default ephemeral key intentionally prevents cross-run correlation.
    /// </summary>
    public string? VerificationManifestPath { get; }
    /// <summary>Maximum mismatch/failure details retained in memory.</summary>
    public int MaxVerificationIssues { get; }
}
