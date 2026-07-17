using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Native directory layout produced by a store export.</summary>
public enum EmailStoreNativeDirectoryFormat {
    /// <summary>Maildir folders containing tmp, new, and cur directories.</summary>
    Maildir = 0,
    /// <summary>Apple Mail EMLX files arranged by source folder.</summary>
    Emlx = 1
}

/// <summary>Selection and failure policy for Maildir or EMLX directory export.</summary>
public sealed class EmailStoreNativeDirectoryExportOptions {
    /// <summary>Creates native directory export options.</summary>
    public EmailStoreNativeDirectoryExportOptions(EmailStoreNativeDirectoryFormat format,
        string? folderId = null, bool includeDescendants = false, bool includeAssociatedItems = false,
        bool includeOrphanedItems = false, bool preserveFolderHierarchy = true,
        bool overwriteExisting = false, bool continueOnError = true, bool writeManifest = true,
        int maxItems = 100_000, EmailWriterOptions? messageOptions = null,
        EmailStoreEmlxWriterOptions? emlxOptions = null) {
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        Format = format;
        FolderId = string.IsNullOrWhiteSpace(folderId) ? null : folderId;
        IncludeDescendants = includeDescendants;
        IncludeAssociatedItems = includeAssociatedItems;
        IncludeOrphanedItems = includeOrphanedItems;
        PreserveFolderHierarchy = preserveFolderHierarchy;
        OverwriteExisting = overwriteExisting;
        ContinueOnError = continueOnError;
        WriteManifest = writeManifest;
        MaxItems = maxItems;
        MessageOptions = messageOptions ?? EmailWriterOptions.Default;
        EmlxOptions = emlxOptions ?? new EmailStoreEmlxWriterOptions(MessageOptions);
    }

    /// <summary>Destination native directory format.</summary>
    public EmailStoreNativeDirectoryFormat Format { get; }
    /// <summary>Optional source folder identifier.</summary>
    public string? FolderId { get; }
    /// <summary>Whether descendant source folders are included.</summary>
    public bool IncludeDescendants { get; }
    /// <summary>Whether folder-associated items are included.</summary>
    public bool IncludeAssociatedItems { get; }
    /// <summary>Whether recovered orphaned items are included.</summary>
    public bool IncludeOrphanedItems { get; }
    /// <summary>Whether the source hierarchy is represented in the destination.</summary>
    public bool PreserveFolderHierarchy { get; }
    /// <summary>Whether existing destination files may be replaced.</summary>
    public bool OverwriteExisting { get; }
    /// <summary>Whether the export continues after an item failure.</summary>
    public bool ContinueOnError { get; }
    /// <summary>Whether a preservation manifest is written.</summary>
    public bool WriteManifest { get; }
    /// <summary>Maximum attempted items.</summary>
    public int MaxItems { get; }
    /// <summary>EML policy used for Maildir messages.</summary>
    public EmailWriterOptions MessageOptions { get; }
    /// <summary>EMLX policy used for Apple Mail artifacts.</summary>
    public EmailStoreEmlxWriterOptions EmlxOptions { get; }
}
