using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Sets properties on a table of contents in a Word document.</summary>
/// <example>
///   <summary>Customize table of contents text during composition.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\Report.docx {
///     Add-OfficeWordTableOfContents
///     Set-OfficeWordTableOfContents -Text 'Contents' -TextNoContent 'No entries yet'
///     Add-OfficeWordParagraph -Text 'Executive summary' -Style Heading1
///     Update-OfficeWordTableOfContents
/// }</code>
///   <para>Updates TOC display text, adds heading content, and marks the TOC for refresh.</para>
/// </example>
/// <example>
///   <summary>Update an existing TOC object from a document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $doc |
///     Get-OfficeWordTableOfContents |
///     Set-OfficeWordTableOfContents -Text 'Report contents'
/// $doc | Save-OfficeWord -Path .\Report-Toc.docx</code>
///   <para>Pipes the OfficeIMO TOC object into the thin setter and saves the document.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeWordTableOfContents")]
[OutputType(typeof(WordTableOfContent))]
public sealed class SetOfficeWordTableOfContentsCommand : PSCmdlet
{
    private const string ParameterSetTableOfContents = "TableOfContents";
    private const string ParameterSetDocument = "Document";

    /// <summary>Table of contents to update.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetTableOfContents)]
    public WordTableOfContent? TableOfContents { get; set; }

    /// <summary>Document to update when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument? Document { get; set; }

    /// <summary>Heading text for the table of contents.</summary>
    [Parameter]
    public string? Text { get; set; }

    /// <summary>Text shown when the table of contents has no entries.</summary>
    [Parameter]
    public string? TextNoContent { get; set; }

    /// <summary>Emit the updated table of contents.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Text == null && TextNoContent == null)
        {
            throw new PSArgumentException("Specify -Text or -TextNoContent to update the table of contents.");
        }

        var toc = ResolveTableOfContents();

        if (Text != null)
        {
            toc.Text = Text;
        }

        if (TextNoContent != null)
        {
            toc.TextNoContent = TextNoContent;
        }

        if (PassThru.IsPresent)
        {
            WriteObject(toc);
        }
    }

    private WordTableOfContent ResolveTableOfContents()
    {
        if (TableOfContents != null)
        {
            return TableOfContents;
        }

        var document = Document ?? WordDslContext.Require(this).Document;
        if (document == null)
        {
            throw new InvalidOperationException("Word document was not provided.");
        }

        var toc = document.TableOfContent;
        if (toc == null)
        {
            throw new InvalidOperationException("Table of contents was not found in the document.");
        }

        return toc;
    }
}
