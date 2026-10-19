using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a repeating section content control to the current paragraph.</summary>
/// <example>
///   <summary>Add a repeating section.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordRepeatingSection -SectionTitle 'Items' -Alias 'LineItems' }</code>
///   <para>Creates a repeating section control.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordRepeatingSection")]
[Alias("WordRepeatingSection")]
[OutputType(typeof(WordRepeatingSection))]
public sealed class AddOfficeWordRepeatingSectionCommand : PSCmdlet
{
    /// <summary>Optional title for the repeating section.</summary>
    [Parameter]
    public string? SectionTitle { get; set; }

    /// <summary>Optional alias for the control.</summary>
    [Parameter]
    public string? Alias { get; set; }

    /// <summary>Optional tag for the control.</summary>
    [Parameter]
    public string? Tag { get; set; }

    /// <summary>Explicit paragraph to receive the control.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Emit the created control.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var paragraph = ResolveParagraph();
        var control = paragraph.AddRepeatingSection(SectionTitle, Alias, Tag);

        if (PassThru.IsPresent)
        {
            WriteObject(control);
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
