using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Email;
using OfficeIMO.Email.Store;

namespace PSWriteOffice.Cmdlets.Email;

/// <summary>Reads a native EML, EMLX, MSG, or TNEF artifact with bounded diagnostics.</summary>
/// <example>
///   <summary>Read a message without retaining attachment payloads.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeEmailReaderOptions -ExcludeAttachmentContent -MaxAttachmentBytes 25MB
/// Get-OfficeEmail -Path .\Message.msg -Options $options -AsResult</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeEmail")]
[OutputType(typeof(EmailDocument), typeof(EmailReadResult), typeof(EmailStoreReadResult))]
public sealed class GetOfficeEmailCommand : PSCmdlet
{
    /// <summary>Path to an EML, EMLX, MSG, TNEF, or winmail.dat file.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional format detection, compound-file, MIME, attachment, and size limits.</summary>
    [Parameter]
    public EmailReaderOptions? Options { get; set; }

    /// <summary>Optional Apple Mail EMLX envelope, metadata, attachment, and size limits.</summary>
    [Parameter]
    public EmailStoreReaderOptions? StoreOptions { get; set; }

    /// <summary>Return the read result with diagnostics and consumed byte count.</summary>
    [Parameter]
    public SwitchParameter AsResult { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (string.Equals(System.IO.Path.GetExtension(input), ".emlx", StringComparison.OrdinalIgnoreCase))
        {
            var storeResult = new EmailStoreReader(StoreOptions ?? CreateStoreOptions()).Read(input);
            if (AsResult.IsPresent)
            {
                WriteObject(storeResult);
                return;
            }

            var documents = storeResult.Store.Folders.SelectMany(folder => folder.Items).Select(item => item.Document).ToArray();
            if (documents.Length != 1)
            {
                throw new InvalidDataException($"The EMLX artifact contained {documents.Length} messages; exactly one was expected.");
            }

            WriteObject(documents[0]);
            return;
        }

        if (StoreOptions != null)
        {
            throw new PSArgumentException("-StoreOptions is supported only for Apple Mail .emlx artifacts.", nameof(StoreOptions));
        }

        var result = new EmailDocumentReader(Options ?? EmailReaderOptions.Default).Read(input);
        WriteObject(AsResult.IsPresent ? result : result.Document);
    }

    private EmailStoreReaderOptions CreateStoreOptions()
    {
        var options = Options ?? EmailReaderOptions.Default;
        return new EmailStoreReaderOptions(
            maxInputBytes: options.MaxInputBytes,
            maxItemCount: 1,
            maxPropertiesPerItem: options.MaxMapiPropertyCount,
            maxDecodedPropertyBytesPerItem: options.MaxDecodedPropertyBytes,
            maxAttachmentsPerItem: options.MaxAttachmentCount,
            maxAttachmentBytes: options.MaxAttachmentBytes,
            maxTotalAttachmentBytes: options.MaxTotalAttachmentBytes,
            retainAttachmentContent: options.IncludeAttachmentContent,
            maxNestedMessageDepth: options.MaxNestedMessageDepth,
            maxMessageBytes: options.MaxInputBytes);
    }
}
