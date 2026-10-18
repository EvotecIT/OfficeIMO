using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a structured content control to the current paragraph.</summary>
/// <example>
///   <summary>Add a content control with text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordContentControl -Text 'Client' -Alias 'ClientName' }</code>
///   <para>Creates a content control with the specified text.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordContentControl")]
[Alias("WordContentControl")]
[OutputType(typeof(WordStructuredDocumentTag))]
public sealed class AddOfficeWordContentControlCommand : PSCmdlet
{
    /// <summary>Initial text for the control.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

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
        var control = paragraph.AddStructuredDocumentTag(Text ?? string.Empty, Alias, Tag);

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
