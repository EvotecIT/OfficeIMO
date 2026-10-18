using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a table of contents to a Word document.</summary>
/// <example>
///   <summary>Add a table of contents before report sections.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\ExecutiveReport.docx {
///     Add-OfficeWordTableOfContents -Style Template1
///     Add-OfficeWordParagraph -Text 'Executive summary' -Style Heading1
///     Add-OfficeWordParagraph -Text 'Summary text'
///     Add-OfficeWordParagraph -Text 'Appendix' -Style Heading1
///     Add-OfficeWordParagraph -Text 'Supporting details'
///     Update-OfficeWordTableOfContents
/// }</code>
///   <para>Creates a navigable report outline and marks the TOC for refresh when the document opens.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordTableOfContents")]
[Alias("WordTableOfContents")]
[OutputType(typeof(WordTableOfContent))]
public sealed class AddOfficeWordTableOfContentsCommand : PSCmdlet
{
    /// <summary>Document to modify when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <summary>Table of contents template style.</summary>
    [Parameter]
    public WordTableOfContentsStyle Style { get; set; } = WordTableOfContentsStyle.Template1;

    /// <summary>Emit the created table of contents.</summary>
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

        var toc = document.AddTableOfContent(Style);

        if (PassThru.IsPresent)
        {
            WriteObject(toc);
        }
    }
}
