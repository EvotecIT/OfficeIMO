using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Discovery, preservation, and failure policy for recoverable-index item export.</summary>
public sealed class EmailStoreRecoveryExportOptions {
    /// <summary>Creates recovery-export options.</summary>
    public EmailStoreRecoveryExportOptions(
        EmailStoreRecoveryOptions? discoveryOptions = null,
        EmailFileFormat format = EmailFileFormat.OutlookMsg,
        bool preserveFolderHierarchy = true,
        bool overwriteExisting = false,
        bool continueOnItemError = true,
        bool writeManifest = true,
        EmailWriterOptions? writerOptions = null) {
        if (format != EmailFileFormat.Eml && format != EmailFileFormat.OutlookMsg &&
            format != EmailFileFormat.OutlookTemplate && format != EmailFileFormat.Tnef) {
            throw new ArgumentException(
                "Recovery export supports EML, MSG, OFT, or TNEF output.", nameof(format));
        }
        DiscoveryOptions = discoveryOptions ?? new EmailStoreRecoveryOptions();
        Format = format;
        PreserveFolderHierarchy = preserveFolderHierarchy;
        OverwriteExisting = overwriteExisting;
        ContinueOnItemError = continueOnItemError;
        WriteManifest = writeManifest;
        WriterOptions = writerOptions ?? EmailWriterOptions.Default;
    }

    /// <summary>Bounds and folder scope used to discover indexed orphans.</summary>
    public EmailStoreRecoveryOptions DiscoveryOptions { get; }
    /// <summary>Per-item preservation format. MSG is the default to retain Outlook/MAPI semantics.</summary>
    public EmailFileFormat Format { get; }
    /// <summary>Whether original folder evidence is represented in destination directories.</summary>
    public bool PreserveFolderHierarchy { get; }
    /// <summary>Whether an existing artifact or manifest may be replaced.</summary>
    public bool OverwriteExisting { get; }
    /// <summary>Whether isolated unreadable items are diagnosed and skipped.</summary>
    public bool ContinueOnItemError { get; }
    /// <summary>Whether a recovery manifest with source folder and item evidence is written.</summary>
    public bool WriteManifest { get; }
    /// <summary>Underlying item writer bounds and conversion policy.</summary>
    public EmailWriterOptions WriterOptions { get; }
}
