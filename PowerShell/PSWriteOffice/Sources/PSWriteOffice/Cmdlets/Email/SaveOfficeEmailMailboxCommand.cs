using System.IO;
using System.Management.Automation;
using OfficeIMO.Email;

namespace PSWriteOffice.Cmdlets.Email;

/// <summary>Saves a native mbox mailbox with output diagnostics.</summary>
/// <example>
///   <summary>Save an mboxrd mailbox and return its diagnostics.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeEmailMailboxWriterOptions -Variant Mboxrd
/// $mailbox | Save-OfficeEmailMailbox -Path .\Archive.mbox -Options $options -PassThru</code>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeEmailMailbox", SupportsShouldProcess = true)]
[OutputType(typeof(EmailWriteResult))]
public sealed class SaveOfficeEmailMailboxCommand : OfficeMutationCmdlet {
    /// <summary>Mailbox to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public EmailMailbox Mailbox { get; set; } = null!;

    /// <summary>Destination mbox path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional mailbox variant, envelope, and output limits.</summary>
    [Parameter]
    public EmailMailboxWriterOptions? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (!ShouldProcess(output, "Save mbox mailbox")) return;
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        WritePassThru(Mailbox.Save(output, Options));
    }
}
