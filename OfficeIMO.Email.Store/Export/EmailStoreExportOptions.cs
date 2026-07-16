using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Selection, format, and failure policy for item-by-item directory export.</summary>
public sealed class EmailStoreExportOptions {
    /// <summary>Creates export options.</summary>
    public EmailStoreExportOptions(
        EmailFileFormat format = EmailFileFormat.Eml,
        string? folderId = null,
        bool includeDescendants = false,
        bool includeAssociatedItems = false,
        bool includeOrphanedItems = false,
        bool preserveFolderHierarchy = true,
        bool overwriteExisting = false,
        bool continueOnError = true,
        bool writeManifest = true,
        int maxItems = 100_000,
        EmailWriterOptions? writerOptions = null) {
        if (format != EmailFileFormat.Eml &&
            format != EmailFileFormat.OutlookMsg &&
            format != EmailFileFormat.OutlookTemplate &&
            format != EmailFileFormat.Tnef) {
            throw new ArgumentException(
                "Store directory export supports EML, MSG, OFT, or TNEF output.", nameof(format));
        }
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
        WriterOptions = writerOptions ?? EmailWriterOptions.Default;
    }

    /// <summary>Destination artifact format.</summary>
    public EmailFileFormat Format { get; }

    /// <summary>Optional source folder identifier. Null exports every folder.</summary>
    public string? FolderId { get; }

    /// <summary>Whether descendants of <see cref="FolderId"/> are included.</summary>
    public bool IncludeDescendants { get; }

    /// <summary>Whether folder-associated information items are exported.</summary>
    public bool IncludeAssociatedItems { get; }

    /// <summary>Whether recoverable items absent from contents tables are exported.</summary>
    public bool IncludeOrphanedItems { get; }

    /// <summary>Whether the source folder hierarchy is represented in destination directories.</summary>
    public bool PreserveFolderHierarchy { get; }

    /// <summary>Whether an existing artifact with the same stable source identifier may be replaced.</summary>
    public bool OverwriteExisting { get; }

    /// <summary>Whether export continues after an item read or write failure.</summary>
    public bool ContinueOnError { get; }

    /// <summary>Whether a tab-separated preservation manifest is written after the export.</summary>
    public bool WriteManifest { get; }

    /// <summary>Maximum item references attempted.</summary>
    public int MaxItems { get; }

    /// <summary>Underlying OfficeIMO.Email conversion and output policy.</summary>
    public EmailWriterOptions WriterOptions { get; }
}
