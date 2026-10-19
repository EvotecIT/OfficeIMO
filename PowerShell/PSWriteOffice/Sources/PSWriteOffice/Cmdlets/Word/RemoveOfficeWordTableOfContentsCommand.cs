using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Removes the table of contents from a Word document.</summary>
/// <example>
///   <summary>Remove a table of contents from an opened document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $doc | Remove-OfficeWordTableOfContents -PassThru |
///     Save-OfficeWord -Path .\Report-NoToc.docx</code>
///   <para>Removes the TOC from an OfficeIMO document object and saves the changed document to a new file.</para>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficeWordTableOfContents")]
public sealed class RemoveOfficeWordTableOfContentsCommand : PSCmdlet
{
    /// <summary>Document to modify when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <summary>Emit the document after removal.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = Document ?? WordDslContext.Require(this).Document;
        if (document == null)
        {
            throw new InvalidOperationException("Word document was not provided.");
        }

        document.RemoveTableOfContent();

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
