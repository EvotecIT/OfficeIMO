using System.Management.Automation;
using OfficeIMO.Email;

namespace PSWriteOffice.Cmdlets.Email;

/// <summary>Creates deterministic mbox writer settings through ordinary PowerShell parameters.</summary>
/// <example>
///   <summary>Write an mboxo mailbox with a reusable per-message policy.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$messageOptions = New-OfficeEmailWriterOptions -IncludeBccHeader
/// $options = New-OfficeEmailMailboxWriterOptions -MessageOptions $messageOptions -Variant Mboxo
/// $mailbox | Save-OfficeEmailMailbox -Path .\Archive.mbox -Options $options -PassThru</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeEmailMailboxWriterOptions")]
[OutputType(typeof(EmailMailboxWriterOptions))]
public sealed class NewOfficeEmailMailboxWriterOptionsCommand : PSCmdlet {
    /// <summary>Serialization policy applied independently to each message.</summary>
    [Parameter(ValueFromPipeline = true)] public EmailWriterOptions? MessageOptions { get; set; }
    /// <summary>Concrete mbox escaping convention to write.</summary>
    [Parameter] public MboxVariant Variant { get; set; } = MboxVariant.Mboxrd;

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(new EmailMailboxWriterOptions(MessageOptions, Variant));
}
