using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds an endnote reference to a Word paragraph.</summary>
/// <example>
///   <summary>Add an endnote inside the Word DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph -Text 'Appendix reference' { Add-OfficeWordEndnote -Text 'Full calculation appears in the appendix.' }</code>
///   <para>Creates an endnote reference on the current paragraph.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordEndnote")]
[Alias("WordEndnote")]
[OutputType(typeof(WordParagraph))]
public sealed class AddOfficeWordEndnoteCommand : PSCmdlet
{
    /// <summary>Endnote text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Paragraph to receive the endnote reference.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Emit the created endnote paragraph.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (string.IsNullOrWhiteSpace(Text))
        {
            throw new PSArgumentException("Endnote text cannot be empty.", nameof(Text));
        }

        var paragraph = ResolveParagraph();
        var note = paragraph.AddEndNote(Text.Trim());

        if (PassThru.IsPresent)
        {
            WriteObject(note);
        }
    }

    private WordParagraph ResolveParagraph()
    {
        if (Paragraph != null)
        {
            return Paragraph;
        }

        var context = WordDslContext.Require(this);
        return context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
    }
}
