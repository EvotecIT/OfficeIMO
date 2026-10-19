using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a bookmark to the current paragraph.</summary>
/// <example>
///   <summary>Add a bookmark in a paragraph.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordText -Text 'Intro'; Add-OfficeWordBookmark -Name 'Intro' }</code>
///   <para>Creates a bookmark named <c>Intro</c> on the paragraph.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordBookmark")]
[Alias("WordBookmark")]
public sealed class AddOfficeWordBookmarkCommand : PSCmdlet
{
    /// <summary>Bookmark name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Explicit paragraph to receive the bookmark.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Emit the paragraph after adding the bookmark.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var paragraph = Paragraph;
        if (paragraph == null)
        {
            var context = WordDslContext.Require(this);
            paragraph = context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
        }

        if (string.IsNullOrWhiteSpace(Name))
        {
            throw new PSArgumentException("Bookmark name cannot be empty.", nameof(Name));
        }

        paragraph.AddBookmark(Name.Trim());

        if (PassThru.IsPresent)
        {
            WriteObject(paragraph);
        }
    }
}
