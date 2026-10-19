using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a date picker content control to the current paragraph.</summary>
/// <example>
///   <summary>Add a date picker with today's date.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordDatePicker -Date (Get-Date) -Alias 'DueDate' }</code>
///   <para>Creates a date picker control with an initial value.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordDatePicker")]
[Alias("WordDatePicker")]
[OutputType(typeof(WordDatePicker))]
public sealed class AddOfficeWordDatePickerCommand : PSCmdlet
{
    /// <summary>Optional initial date.</summary>
    [Parameter]
    public DateTime? Date { get; set; }

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
        var control = paragraph.AddDatePicker(Date, Alias, Tag);

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
