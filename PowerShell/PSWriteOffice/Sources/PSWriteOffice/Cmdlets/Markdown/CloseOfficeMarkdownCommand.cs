using System.Management.Automation;
using OfficeIMO.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Completes a live Markdown document's lifetime without saving.</summary>
/// <example><summary>Discard pending changes.</summary>
/// <code>$document | Close-OfficeMarkdown</code></example>
[Cmdlet(VerbsCommon.Close, "OfficeMarkdown")]
public sealed class CloseOfficeMarkdownCommand : PSCmdlet {
    /// <summary>Markdown document to close.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true), ValidateNotNull]
    public MarkdownDocument Document { get; set; } = null!;
    /// <inheritdoc />
    protected override void ProcessRecord() => Document.Dispose();
}
