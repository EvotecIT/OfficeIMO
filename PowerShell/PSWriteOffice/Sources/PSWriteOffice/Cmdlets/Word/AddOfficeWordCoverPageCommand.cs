using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a built-in cover page template to a Word document.</summary>
/// <para>Uses OfficeIMO.Word cover page templates and optional cover-page metadata.</para>
/// <example>
///   <summary>Add a cover page in the Word DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\Report.docx { Add-OfficeWordCoverPage -Template Element -Abstract 'Executive summary' }</code>
///   <para>Creates a document with a template-driven cover page.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordCoverPage")]
[Alias("WordCoverPage")]
[OutputType(typeof(WordCoverPage))]
public sealed class AddOfficeWordCoverPageCommand : PSCmdlet
{
    /// <summary>Cover page template to insert.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public WordCoverPageTemplate Template { get; set; }

    /// <summary>Document to update. Defaults to the current Word DSL or tracked document.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <summary>Publish date stored in the cover page properties custom XML part.</summary>
    [Parameter]
    public string? PublishDate { get; set; }

    /// <summary>Abstract/summary stored in the cover page properties custom XML part.</summary>
    [Parameter]
    public string? Abstract { get; set; }

    /// <summary>Company address stored in the cover page properties custom XML part.</summary>
    [Parameter]
    public string? CompanyAddress { get; set; }

    /// <summary>Company email stored in the cover page properties custom XML part.</summary>
    [Parameter]
    public string? CompanyEmail { get; set; }

    /// <summary>Emit the created cover page.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = ResolveDocument();
        ApplyProperties(document);
        var coverPage = document.AddCoverPage(Template);

        if (PassThru.IsPresent)
        {
            WriteObject(coverPage);
        }
    }

    private WordDocument ResolveDocument()
    {
        if (Document != null)
        {
            return Document;
        }

        var context = WordDslContext.Current;
        if (context != null)
        {
            return context.Document;
        }

        return WordDocumentService.GetCurrentTrackedDocument()
            ?? throw new InvalidOperationException("No active Word document was found. Pass -Document or call this inside New-OfficeWord.");
    }

    private void ApplyProperties(WordDocument document)
    {
        if (PublishDate != null)
        {
            document.CoverPageProperties.PublishDate = PublishDate;
        }
        if (Abstract != null)
        {
            document.CoverPageProperties.Abstract = Abstract;
        }
        if (CompanyAddress != null)
        {
            document.CoverPageProperties.CompanyAddress = CompanyAddress;
        }
        if (CompanyEmail != null)
        {
            document.CoverPageProperties.CompanyEmail = CompanyEmail;
        }
    }
}
