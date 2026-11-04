using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using System.Management.Automation;
using OfficeIMO.Email;

namespace PSWriteOffice.Cmdlets.Email;

/// <summary>Reads a native mbox mailbox with bounded per-message diagnostics.</summary>
/// <example>
///   <summary>Read a bounded mbox mailbox with diagnostics.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeEmailMailboxReaderOptions -MaxMessageCount 5000
/// Open-OfficeEmailMailbox -Path .\Archive.mbox -Options $options -AsResult</code>
/// </example>
[Cmdlet(VerbsCommon.Open, "OfficeEmailMailbox", DefaultParameterSetName = "Path")]
[OutputType(typeof(EmailMailbox), typeof(EmailMailboxReadResult))]
public sealed class OpenOfficeEmailMailboxCommand : PSCmdlet
{
    /// <summary>Path to an mbox or mbx file.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Exact input path without wildcard expansion.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    public string LiteralPath { get; set; } = string.Empty;


    /// <summary>Optional mailbox variant and input limits.</summary>
    [Parameter]
    public EmailMailboxReaderOptions? Options { get; set; }

    /// <summary>Return the mailbox read result with diagnostics.</summary>
    [Parameter]
    public SwitchParameter AsResult { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var path in OfficePath.ResolveInput(SessionState, ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath", allowDirectory: false)) ReadPath(path);
    }
    private void ReadPath(string inputPath) {
        var result = new EmailMailboxReader(Options ?? EmailMailboxReaderOptions.Default).Read(
            inputPath);
        WriteObject(AsResult.IsPresent ? result : result.Mailbox);
    }
}
