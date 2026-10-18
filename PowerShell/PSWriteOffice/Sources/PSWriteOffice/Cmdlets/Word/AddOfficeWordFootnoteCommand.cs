using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a footnote reference to a Word paragraph.</summary>
/// <example>
///   <summary>Add a footnote inside the Word DSL.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph -Text 'Service availability' { Add-OfficeWordFootnote -Text 'Measured from successful health probes.' }</code>
///   <para>Creates a footnote reference on the current paragraph.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordFootnote")]
[Alias("WordFootnote")]
[OutputType(typeof(WordParagraph))]
public sealed class AddOfficeWordFootnoteCommand : PSCmdlet
{
    /// <summary>Footnote text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Paragraph to receive the footnote reference.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Emit the created footnote paragraph.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (string.IsNullOrWhiteSpace(Text))
        {
            throw new PSArgumentException("Footnote text cannot be empty.", nameof(Text));
        }

        var paragraph = ResolveParagraph();
        var note = paragraph.AddFootNote(Text.Trim());

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
