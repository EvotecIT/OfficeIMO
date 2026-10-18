using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a checkbox content control to the current paragraph.</summary>
/// <example>
///   <summary>Add a checked checkbox.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordCheckBox -Checked -Alias 'Approved' }</code>
///   <para>Creates a checked checkbox content control.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordCheckBox")]
[Alias("WordCheckBox")]
[OutputType(typeof(WordCheckBox))]
public sealed class AddOfficeWordCheckBoxCommand : PSCmdlet
{
    /// <summary>Set the checkbox as checked.</summary>
    [Parameter]
    public SwitchParameter Checked { get; set; }

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
        var control = paragraph.AddCheckBox(Checked.IsPresent, Alias, Tag);

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
