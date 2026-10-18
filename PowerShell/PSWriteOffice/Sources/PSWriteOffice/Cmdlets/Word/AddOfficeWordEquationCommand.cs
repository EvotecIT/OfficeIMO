using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds an Office Math equation to a Word document or paragraph.</summary>
/// <para>Accepts OMML and keeps conversion/parsing outside the cmdlet.</para>
/// <example>
///   <summary>Add an equation to a generated report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$omml = '&lt;m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"&gt;&lt;m:r&gt;&lt;m:t&gt;x+1&lt;/m:t&gt;&lt;/m:r&gt;&lt;/m:oMath&gt;'
/// New-OfficeWord -Path .\Formula.docx {
///     Add-OfficeWordParagraph -Text 'The following expression is stored as Office Math.'
///     Add-OfficeWordEquation -Omml $omml
/// }</code>
///   <para>Inserts prebuilt OMML into the current document; conversion to OMML is intentionally outside the cmdlet.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordEquation")]
[Alias("WordEquation")]
[OutputType(typeof(WordParagraph))]
public sealed class AddOfficeWordEquationCommand : PSCmdlet
{
    /// <summary>Office Math Markup Language content.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Omml { get; set; } = string.Empty;

    /// <summary>Paragraph to receive the equation.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Document to receive a new equation paragraph.</summary>
    [Parameter]
    public WordDocument? Document { get; set; }

    /// <summary>Emit the equation paragraph.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (string.IsNullOrWhiteSpace(Omml))
        {
            throw new PSArgumentException("Equation OMML cannot be empty.", nameof(Omml));
        }

        var paragraph = ResolveParagraph();
        var result = paragraph.AddEquation(Omml);

        if (PassThru.IsPresent)
        {
            WriteObject(result);
        }
    }

    private WordParagraph ResolveParagraph()
    {
        if (Paragraph != null)
        {
            return Paragraph;
        }

        if (Document != null)
        {
            return Document.AddParagraph();
        }

        var context = WordDslContext.Current;
        if (context != null)
        {
            return context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
        }

        var document = WordDocumentService.GetCurrentTrackedDocument()
            ?? throw new InvalidOperationException("No active Word document was found. Pass -Document, pipe a paragraph, or call this inside New-OfficeWord.");
        return document.AddParagraph();
    }
}
