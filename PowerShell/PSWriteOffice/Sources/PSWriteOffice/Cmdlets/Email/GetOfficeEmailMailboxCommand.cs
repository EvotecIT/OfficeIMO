using System.Management.Automation;
using OfficeIMO.Email;

namespace PSWriteOffice.Cmdlets.Email;

/// <summary>Reads a native mbox mailbox with bounded per-message diagnostics.</summary>
/// <example>
///   <summary>Read a bounded mbox mailbox with diagnostics.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeEmailMailboxReaderOptions -MaxMessageCount 5000
/// Get-OfficeEmailMailbox -Path .\Archive.mbox -Options $options -AsResult</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeEmailMailbox")]
[OutputType(typeof(EmailMailbox), typeof(EmailMailboxReadResult))]
public sealed class GetOfficeEmailMailboxCommand : PSCmdlet
{
    /// <summary>Path to an mbox or mbx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional mailbox variant and input limits.</summary>
    [Parameter]
    public EmailMailboxReaderOptions? Options { get; set; }

    /// <summary>Return the mailbox read result with diagnostics.</summary>
    [Parameter]
    public SwitchParameter AsResult { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var result = new EmailMailboxReader(Options ?? EmailMailboxReaderOptions.Default).Read(
            SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path));
        WriteObject(AsResult.IsPresent ? result : result.Mailbox);
    }
}
