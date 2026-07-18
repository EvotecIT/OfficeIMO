using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Controls a verified atomic rewrite used to mutate an existing Unicode PST.</summary>
public sealed class EmailStorePstMutationOptions {
    /// <summary>Creates bounded mutation options.</summary>
    public EmailStorePstMutationOptions(
        int maxFolderCount = 100_000,
        int maxItemCount = 1_000_000,
        int maxNestedMessageDepth = 32,
        bool failOnDataLoss = true,
        bool verifyAfterWrite = true,
        EmailSemanticComparisonOptions? verificationOptions = null,
        int maxVerificationIssues = 1_000,
        string? backupPath = null,
        bool overwriteBackup = false,
        string? pstPassword = null,
        Encoding? pstPasswordEncoding = null) {
        if (maxFolderCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxFolderCount));
        if (maxItemCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxItemCount));
        if (maxNestedMessageDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxNestedMessageDepth));
        if (maxVerificationIssues <= 0) throw new ArgumentOutOfRangeException(nameof(maxVerificationIssues));
        MaxFolderCount = maxFolderCount;
        MaxItemCount = maxItemCount;
        MaxNestedMessageDepth = maxNestedMessageDepth;
        FailOnDataLoss = failOnDataLoss;
        VerifyAfterWrite = verifyAfterWrite;
        VerificationOptions = verificationOptions;
        MaxVerificationIssues = maxVerificationIssues;
        BackupPath = string.IsNullOrWhiteSpace(backupPath) ? null : Path.GetFullPath(backupPath);
        OverwriteBackup = overwriteBackup;
        PstPassword = pstPassword;
        PstPasswordEncoding = pstPasswordEncoding ?? Encoding.ASCII;
    }

    /// <summary>Maximum source and resulting folders accepted by the transaction.</summary>
    public int MaxFolderCount { get; }

    /// <summary>Maximum source and resulting items accepted by the transaction.</summary>
    public int MaxItemCount { get; }

    /// <summary>Maximum embedded-message nesting depth read, written, and compared.</summary>
    public int MaxNestedMessageDepth { get; }

    /// <summary>Whether any fidelity warning or error prevents replacement of the original PST.</summary>
    public bool FailOnDataLoss { get; }

    /// <summary>Whether every resulting item is reopened and semantically compared before replacement.</summary>
    public bool VerifyAfterWrite { get; }

    /// <summary>Optional semantic comparison policy used by post-write verification.</summary>
    public EmailSemanticComparisonOptions? VerificationOptions { get; }

    /// <summary>Maximum mismatch and failure details retained in memory.</summary>
    public int MaxVerificationIssues { get; }

    /// <summary>Optional path that receives a byte-for-byte backup before the original is replaced.</summary>
    public string? BackupPath { get; }

    /// <summary>Whether an existing backup path may be atomically replaced.</summary>
    public bool OverwriteBackup { get; }

    /// <summary>Legacy PST password used to validate the source, when required.</summary>
    public string? PstPassword { get; }

    /// <summary>Encoding used for legacy PST password validation.</summary>
    public Encoding PstPasswordEncoding { get; }
}
