namespace OfficeIMO.Email.Store;

/// <summary>Constant-memory catalog and declared-count snapshot of an open email store.</summary>
public sealed class EmailStoreInspectionReport {
    internal EmailStoreInspectionReport(EmailStoreSession session) {
        Format = session.Format;
        DisplayName = session.DisplayName;
        SourceLength = session.SourceLength;
        FolderCount = session.Folders.Count;
        long itemCount = 0;
        long associatedItemCount = 0;
        int unknownItemFolders = 0;
        int unknownAssociatedFolders = 0;
        foreach (EmailStoreFolderInfo folder in session.Folders) {
            if (folder.ItemCount.HasValue) itemCount = checked(itemCount + folder.ItemCount.Value);
            else unknownItemFolders++;
            if (folder.AssociatedItemCount.HasValue) {
                associatedItemCount = checked(associatedItemCount + folder.AssociatedItemCount.Value);
            } else {
                unknownAssociatedFolders++;
            }
        }
        DeclaredItemCount = itemCount;
        DeclaredAssociatedItemCount = associatedItemCount;
        FoldersWithUnknownItemCount = unknownItemFolders;
        FoldersWithUnknownAssociatedItemCount = unknownAssociatedFolders;
        Diagnostics = session.Diagnostics.ToArray();
    }

    /// <summary>Detected store format.</summary>
    public EmailStoreFormat Format { get; }

    /// <summary>Source display name when declared.</summary>
    public string? DisplayName { get; }

    /// <summary>Validated source length.</summary>
    public long SourceLength { get; }

    /// <summary>Number of folders in the lightweight catalog.</summary>
    public int FolderCount { get; }

    /// <summary>Sum of available declared visible-item counts.</summary>
    public long DeclaredItemCount { get; }

    /// <summary>Sum of available declared associated-item counts.</summary>
    public long DeclaredAssociatedItemCount { get; }

    /// <summary>Folders whose source did not expose a visible-item count.</summary>
    public int FoldersWithUnknownItemCount { get; }

    /// <summary>Folders whose source did not expose an associated-item count.</summary>
    public int FoldersWithUnknownAssociatedItemCount { get; }

    /// <summary>Whether every folder exposed a visible-item count.</summary>
    public bool HasCompleteDeclaredItemCount => FoldersWithUnknownItemCount == 0;

    /// <summary>Whether every folder exposed an associated-item count.</summary>
    public bool HasCompleteDeclaredAssociatedItemCount => FoldersWithUnknownAssociatedItemCount == 0;

    /// <summary>Diagnostics emitted while opening and cataloging the store.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>Whether cataloging emitted an error.</summary>
    public bool HasErrors => Diagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error);
}
