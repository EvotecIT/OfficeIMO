using System;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown heading.</summary>
/// <example>
///   <summary>Add headings that feed the Markdown table of contents.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\Report.md {
///     Add-OfficeMarkdownHeading -Level 1 -Text 'Operational Report'
///     Add-OfficeMarkdownTableOfContents -Title 'Contents' -MinLevel 2 -MaxLevel 3 -PlaceAtTop
///     Add-OfficeMarkdownHeading -Level 2 -Text 'Overview'
///     Add-OfficeMarkdownParagraph -Text 'Current operational state.'
/// }</code>
///   <para>Appends headings that can be included in a generated table of contents.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownHeading", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownHeading")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownHeadingCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Heading level (1-6).</summary>
    [Parameter(Position = 0)]
    public int Level { get; set; } = 1;

    /// <summary>Heading text.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Emit the Markdown document after appending the heading.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        if (Level < 1 || Level > 6)
        {
            throw new PSArgumentOutOfRangeException(nameof(Level), Level, "Heading level must be between 1 and 6.");
        }

        switch (Level)
        {
            case 1:
                doc.H1(Text);
                break;
            case 2:
                doc.H2(Text);
                break;
            case 3:
                doc.H3(Text);
                break;
            case 4:
                doc.H4(Text);
                break;
            case 5:
                doc.H5(Text);
                break;
            default:
                doc.H6(Text);
                break;
        }

        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
    }

    private MarkdownDoc ResolveDocument()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            return Document ?? throw new PSArgumentException("Provide a Markdown document.");
        }

        return MarkdownDslContext.Require(this).Document;
    }
}
