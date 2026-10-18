using System.Management.Automation;
using System.Text;
using OfficeIMO.Email.Store;

namespace PSWriteOffice.Cmdlets.Email;

/// <summary>Creates bounded email-store reader settings without requiring .NET constructor syntax.</summary>
/// <example>
///   <summary>Read an EMLX message without retaining attachment payloads.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeEmailStoreReaderOptions -ExcludeAttachmentContent -MaxAttachmentsPerItem 100
/// Get-OfficeEmail -Path .\Message.emlx -StoreOptions $options -AsResult</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeEmailStoreReaderOptions")]
[OutputType(typeof(EmailStoreReaderOptions))]
public sealed class NewOfficeEmailStoreReaderOptionsCommand : PSCmdlet {
    /// <summary>Maximum seekable source length.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxInputBytes { get; set; } = 1L * 1024 * 1024 * 1024 * 1024;
    /// <summary>Maximum NDB nodes and blocks visited.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxNodeCount { get; set; } = 25_000_000;
    /// <summary>Maximum tree traversal depth.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxBTreeDepth { get; set; } = 32;
    /// <summary>Maximum PST/OST B-tree pages retained by the cache.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxCachedBTreePages { get; set; } = 512;
    /// <summary>Maximum folders materialized.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxFolderCount { get; set; } = 100000;
    /// <summary>Maximum items materialized.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxItemCount { get; set; } = 1000000;
    /// <summary>Maximum MAPI properties decoded per item.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxPropertiesPerItem { get; set; } = 16384;
    /// <summary>Maximum decoded property bytes per item.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxDecodedPropertyBytesPerItem { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum attachments per item.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxAttachmentsPerItem { get; set; } = 10000;
    /// <summary>Maximum decoded bytes in one attachment.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxAttachmentBytes { get; set; } = 512L * 1024 * 1024;
    /// <summary>Maximum decoded attachment bytes across the read.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxTotalAttachmentBytes { get; set; } = 4L * 1024 * 1024 * 1024;
    /// <summary>Do not retain decoded attachment payloads in memory.</summary>
    [Parameter] public SwitchParameter ExcludeAttachmentContent { get; set; }
    /// <summary>Password used to validate legacy protected PST files.</summary>
    [Parameter] public string? PstPassword { get; set; }
    /// <summary>Encoding name used for the legacy PST password checksum.</summary>
    [Parameter] public string PstPasswordEncoding { get; set; } = "us-ascii";
    /// <summary>Materialize folder-associated information items.</summary>
    [Parameter] public SwitchParameter IncludeAssociatedItems { get; set; }
    /// <summary>Recover item nodes absent from folder contents tables.</summary>
    [Parameter] public SwitchParameter IncludeOrphanedItems { get; set; }
    /// <summary>Maximum embedded-message recursion depth.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int MaxNestedMessageDepth { get; set; } = 16;
    /// <summary>Maximum entries accepted from a compressed email-store archive.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxArchiveEntries { get; set; } = 500000;
    /// <summary>Maximum decoded size declared by one archive entry.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxArchiveEntryBytes { get; set; } = 512L * 1024 * 1024;
    /// <summary>Maximum total decoded size declared by archive entries.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxArchiveDecodedBytes { get; set; } = 8L * 1024 * 1024 * 1024;
    /// <summary>Maximum XML characters parsed from one archive item.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxXmlCharactersPerItem { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum RFC 5322/MIME message bytes accepted from one item.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxMessageBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum directory depth traversed by mailbox-directory sessions.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxDirectoryDepth { get; set; } = 64;
    /// <summary>Maximum EML, EMLX, and Maildir files indexed by one directory session.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxDirectoryFileCount { get; set; } = 1000000;
    /// <summary>Maximum decoded bytes traversed from one PST/OST table data tree.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxDecodedTableBytes { get; set; } = 8L * 1024 * 1024 * 1024;

    /// <inheritdoc />
    protected override void ProcessRecord() {
        Encoding encoding;
        try {
            encoding = Encoding.GetEncoding(PstPasswordEncoding);
        } catch (System.Exception) {
            throw new PSArgumentException($"Unknown PST password encoding '{PstPasswordEncoding}'.", nameof(PstPasswordEncoding));
        }

        WriteObject(new EmailStoreReaderOptions(
            MaxInputBytes,
            MaxNodeCount,
            MaxBTreeDepth,
            MaxCachedBTreePages,
            MaxFolderCount,
            MaxItemCount,
            MaxPropertiesPerItem,
            MaxDecodedPropertyBytesPerItem,
            MaxAttachmentsPerItem,
            MaxAttachmentBytes,
            MaxTotalAttachmentBytes,
            !ExcludeAttachmentContent.IsPresent,
            PstPassword,
            encoding,
            IncludeAssociatedItems.IsPresent,
            IncludeOrphanedItems.IsPresent,
            MaxNestedMessageDepth,
            MaxArchiveEntries,
            MaxArchiveEntryBytes,
            MaxArchiveDecodedBytes,
            MaxXmlCharactersPerItem,
            MaxMessageBytes,
            MaxDirectoryDepth,
            MaxDirectoryFileCount,
            MaxDecodedTableBytes));
    }
}
