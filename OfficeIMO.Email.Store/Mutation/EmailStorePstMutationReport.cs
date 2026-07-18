namespace OfficeIMO.Email.Store;

/// <summary>Outcome of a committed existing-PST mutation transaction.</summary>
public sealed class EmailStorePstMutationReport {
    internal EmailStorePstMutationReport(string sourcePath, string? backupPath,
        EmailStorePstWriteReport writeReport, EmailStorePstMutationVerificationReport? verification,
        int createdFolders, int renamedFolders, int movedFolders, int deletedFolders,
        int addedItems, int replacedItems, int movedItems, int deletedItems,
        IReadOnlyDictionary<string, string> folderIdMap,
        IReadOnlyDictionary<string, string> itemIdMap,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        SourcePath = sourcePath;
        BackupPath = backupPath;
        WriteReport = writeReport;
        Verification = verification;
        CreatedFolders = createdFolders;
        RenamedFolders = renamedFolders;
        MovedFolders = movedFolders;
        DeletedFolders = deletedFolders;
        AddedItems = addedItems;
        ReplacedItems = replacedItems;
        MovedItems = movedItems;
        DeletedItems = deletedItems;
        FolderIdMap = folderIdMap;
        ItemIdMap = itemIdMap;
        Diagnostics = diagnostics;
    }

    /// <summary>Full path of the atomically replaced PST.</summary>
    public string SourcePath { get; }

    /// <summary>Full path of the committed byte-for-byte backup, or null when none was requested.</summary>
    public string? BackupPath { get; }

    /// <summary>Report for the newly built PST.</summary>
    public EmailStorePstWriteReport WriteReport { get; }

    /// <summary>Post-write semantic verification, or null when explicitly disabled.</summary>
    public EmailStorePstMutationVerificationReport? Verification { get; }

    /// <summary>Number of folders created by the transaction.</summary>
    public int CreatedFolders { get; }

    /// <summary>Number of source folders renamed by the transaction.</summary>
    public int RenamedFolders { get; }

    /// <summary>Number of source folders moved by the transaction.</summary>
    public int MovedFolders { get; }

    /// <summary>Number of source folders removed by the transaction, including recursive descendants.</summary>
    public int DeletedFolders { get; }

    /// <summary>Number of items added by the transaction.</summary>
    public int AddedItems { get; }

    /// <summary>Number of source items replaced by the transaction.</summary>
    public int ReplacedItems { get; }

    /// <summary>Number of source items moved or changed between normal and associated contents.</summary>
    public int MovedItems { get; }

    /// <summary>Number of source or newly staged items removed by the transaction.</summary>
    public int DeletedItems { get; }

    /// <summary>Maps retained source and transaction folder identifiers to identifiers in the rewritten PST.</summary>
    public IReadOnlyDictionary<string, string> FolderIdMap { get; }

    /// <summary>Maps retained source and transaction item identifiers to identifiers in the rewritten PST.</summary>
    public IReadOnlyDictionary<string, string> ItemIdMap { get; }

    /// <summary>Combined transaction, writer, and verification diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>True when a warning, error, or semantic mismatch was reported.</summary>
    public bool HasDataLoss => Verification?.IsSuccessful == false || Diagnostics.Any(item =>
        item.Severity != EmailStoreDiagnosticSeverity.Information);
}
