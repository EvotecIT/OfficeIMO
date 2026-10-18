using System.Management.Automation;
using OfficeIMO.Email;

namespace PSWriteOffice.Cmdlets.Email;

/// <summary>Creates bounded mbox reader settings through ordinary PowerShell parameters.</summary>
/// <example>
///   <summary>Read a bounded mailbox with a reusable per-message policy.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$messageOptions = New-OfficeEmailReaderOptions -ExcludeAttachmentContent
/// $options = New-OfficeEmailMailboxReaderOptions -MessageOptions $messageOptions -MaxMessageCount 5000
/// Get-OfficeEmailMailbox -Path .\Archive.mbox -Options $options -AsResult</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeEmailMailboxReaderOptions")]
[OutputType(typeof(EmailMailboxReaderOptions))]
public sealed class NewOfficeEmailMailboxReaderOptionsCommand : PSCmdlet {
    /// <summary>Bounded policy applied independently to each message.</summary>
    [Parameter(ValueFromPipeline = true)] public EmailReaderOptions? MessageOptions { get; set; }
    /// <summary>Escaping convention to decode.</summary>
    [Parameter] public MboxVariant Variant { get; set; } = MboxVariant.Auto;
    /// <summary>Maximum messages in one mailbox.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaxMessageCount { get; set; } = 100000;
    /// <summary>Maximum aggregate source bytes consumed from one mailbox.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long MaxMailboxBytes { get; set; } = 512L * 1024L * 1024L;

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(new EmailMailboxReaderOptions(
        MaxMailboxBytes,
        MessageOptions,
        Variant,
        MaxMessageCount));
}
