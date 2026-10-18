using System;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Adds a hyperlink to the current Word paragraph.</summary>
/// <example>
///   <summary>Add an external hyperlink.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeWordParagraph { Add-OfficeWordHyperlink -Text 'Example' -Url 'https://example.org' -Styled }</code>
///   <para>Creates a styled external hyperlink in the active paragraph.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordHyperlink", DefaultParameterSetName = ParameterSetContextUrl)]
[Alias("WordHyperlink")]
[OutputType(typeof(WordHyperLink))]
public sealed class AddOfficeWordHyperlinkCommand : PSCmdlet
{
    private const string ParameterSetContextUrl = "ContextUrl";
    private const string ParameterSetContextAnchor = "ContextAnchor";
    private const string ParameterSetParagraphUrl = "ParagraphUrl";
    private const string ParameterSetParagraphAnchor = "ParagraphAnchor";

    /// <summary>Paragraph to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetParagraphUrl)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetParagraphAnchor)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Displayed hyperlink text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>External hyperlink URL.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetContextUrl)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetParagraphUrl)]
    [Alias("Uri")]
    public string Url { get; set; } = string.Empty;

    /// <summary>Bookmark anchor target within the document.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetContextAnchor)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetParagraphAnchor)]
    [Alias("Bookmark")]
    public string Anchor { get; set; } = string.Empty;

    /// <summary>Apply the built-in hyperlink style.</summary>
    [Parameter]
    public SwitchParameter Styled { get; set; }

    /// <summary>Optional hyperlink tooltip.</summary>
    [Parameter]
    public string Tooltip { get; set; } = string.Empty;

    /// <summary>Do not mark the hyperlink in navigation history.</summary>
    [Parameter]
    public SwitchParameter NoHistory { get; set; }

    /// <summary>Emit the created hyperlink.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (string.IsNullOrWhiteSpace(Text))
        {
            throw new PSArgumentException("Provide hyperlink text.", nameof(Text));
        }

        var paragraph = Paragraph ?? ResolveParagraph();
        WordParagraph updatedParagraph;

        if (ParameterSetName == ParameterSetContextAnchor || ParameterSetName == ParameterSetParagraphAnchor)
        {
            if (string.IsNullOrWhiteSpace(Anchor))
            {
                throw new PSArgumentException("Provide a bookmark anchor.", nameof(Anchor));
            }

            updatedParagraph = paragraph.AddHyperLink(Text, Anchor, Styled.IsPresent, Tooltip ?? string.Empty, !NoHistory.IsPresent);
        }
        else
        {
            if (!System.Uri.TryCreate(Url, UriKind.Absolute, out var uri))
            {
                throw new PSArgumentException("Provide an absolute URL such as https://example.org or mailto:user@example.org.", nameof(Url));
            }

            updatedParagraph = paragraph.AddHyperLink(Text, uri, Styled.IsPresent, Tooltip ?? string.Empty, !NoHistory.IsPresent);
        }

        if (PassThru.IsPresent)
        {
            var hyperlink = updatedParagraph.Hyperlink;
            if (hyperlink == null)
            {
                throw new InvalidOperationException("Hyperlink was not created.");
            }

            WriteObject(hyperlink);
        }
    }

    private WordParagraph ResolveParagraph()
    {
        var context = WordDslContext.Require(this);
        return context.CurrentParagraph ?? context.RequireParagraphHost().AddParagraph();
    }
}
