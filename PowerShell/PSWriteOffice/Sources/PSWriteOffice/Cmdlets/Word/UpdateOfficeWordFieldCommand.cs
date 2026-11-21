using System.Management.Automation;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Updates fields in a Word document.</summary>
/// <para>Refreshes page number fields and queues table-of-contents updates.</para>
/// <example>
///   <summary>Update fields before saving a generated report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\Report.docx {
///     Add-OfficeWordParagraph -Text 'Release report'
///     Add-OfficeWordTableOfContents
///     Update-OfficeWordField
/// }</code>
///   <para>Updates PAGE and NUMPAGES fields and marks TOC fields as dirty before saving.</para>
/// </example>
/// <example>
///   <summary>Write table-of-contents page numbers from the PDF layout.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$pdf = New-OfficeWordPdfOptions -Font 'Carlito'
/// $document = Open-OfficeWord -Path .\Report.docx
/// $document | Update-OfficeWordField -TableOfContents -LayoutOptions $pdf
/// Close-OfficeWord -Document $document -Save</code>
///   <para>Regenerates the table of contents and refreshes the supported fields, with page numbers from the same layout and fonts as PDF export with $pdf.</para>
/// </example>
[Cmdlet(VerbsData.Update, "OfficeWordField", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordDocument))]
public sealed class UpdateOfficeWordFieldCommand : PSCmdlet
{
    /// <summary>Document to update when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Default")]
    public WordDocument? Document { get; set; }

    /// <summary>Regenerate tables of contents with page numbers from the PDF layout.</summary>
    [Parameter]
    public SwitchParameter TableOfContents { get; set; }

    /// <summary>PDF export options whose layout, fonts and resource policy supply the page numbers of PAGEREF, NUMPAGES and body PAGE fields, including table-of-contents entries. Header and footer PAGE fields keep their stored result and SECTIONPAGES counts explicit breaks. Implies a full refresh of supported fields.</summary>
    [Parameter]
    public WordPdfExportOptions? LayoutOptions { get; set; }

    /// <summary>Emit the updated document.</summary>
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

        if (TableOfContents.IsPresent || LayoutOptions != null)
        {
            var options = new WordFieldUpdateOptions { UpdateTablesOfContents = TableOfContents.IsPresent };
            document.UpdateFields(options.WithPdfLayoutPagination(LayoutOptions ?? new WordPdfExportOptions()));
        }
        else
        {
            document.UpdateFields();
        }

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
