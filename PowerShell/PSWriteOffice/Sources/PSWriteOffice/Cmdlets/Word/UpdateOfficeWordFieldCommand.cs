using System;
using System.Management.Automation;
using OfficeIMO.Word;
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
[Cmdlet(VerbsData.Update, "OfficeWordField", DefaultParameterSetName = "Default")]
[OutputType(typeof(WordDocument))]
public sealed class UpdateOfficeWordFieldCommand : PSCmdlet
{
    /// <summary>Document to update when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = "Default")]
    public WordDocument? Document { get; set; }

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

        document.UpdateFields();

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
