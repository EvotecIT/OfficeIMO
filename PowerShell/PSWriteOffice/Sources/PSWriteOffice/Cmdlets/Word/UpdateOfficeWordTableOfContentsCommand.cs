using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates the table of contents in a Word document.</summary>
/// <example>
///   <summary>Mark table of contents for refresh on open.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\ExecutiveReport.docx {
///     Add-OfficeWordTableOfContents
///     Add-OfficeWordParagraph -Text 'Executive summary' -Style Heading1
///     Add-OfficeWordParagraph -Text 'Summary text'
///     Update-OfficeWordTableOfContents
/// }</code>
///   <para>Marks TOC fields as dirty and updates the document settings so Word refreshes the TOC when opened.</para>
/// </example>
/// <example>
///   <summary>Regenerate a TOC in an existing document object.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $doc | Update-OfficeWordTableOfContents -Regenerate
/// $doc | Save-OfficeWord -Path .\Report-RegeneratedToc.docx</code>
///   <para>Uses OfficeIMO's regenerate path, then saves the updated document.</para>
/// </example>
[Cmdlet(VerbsData.Update, "OfficeWordTableOfContents", DefaultParameterSetName = ParameterSetDocument)]
[OutputType(typeof(WordTableOfContent))]
public sealed class UpdateOfficeWordTableOfContentsCommand : PSCmdlet
{
    private const string ParameterSetTableOfContents = "TableOfContents";
    private const string ParameterSetDocument = "Document";

    /// <summary>Table of contents to update.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetTableOfContents)]
    public WordTableOfContent? TableOfContents { get; set; }

    /// <summary>Document to update when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument? Document { get; set; }

    /// <summary>Rebuild the table of contents before updating.</summary>
    [Parameter]
    public SwitchParameter Regenerate { get; set; }

    /// <summary>Emit the updated table of contents.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var toc = TableOfContents;
        var document = Document;
        if (toc == null && document == null)
        {
            document = WordDslContext.Require(this).Document;
        }

        if (toc == null && document == null)
        {
            throw new InvalidOperationException("Word document was not provided.");
        }

        if (Regenerate.IsPresent)
        {
            toc = toc != null ? toc.Regenerate() : document!.RegenerateTableOfContent();
        }
        else
        {
            if (toc == null)
            {
                toc = document!.TableOfContent;
            }

            if (toc == null)
            {
                throw new InvalidOperationException("Table of contents was not found in the document.");
            }

            toc.Update();
        }

        if (PassThru.IsPresent && toc != null)
        {
            WriteObject(toc);
        }
    }
}
