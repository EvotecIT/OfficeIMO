using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Placement of a generated Markdown table of contents.</summary>
public enum MarkdownTocPlacement {
    /// <summary>Insert where the command runs.</summary>
    Current,
    /// <summary>Insert at the start of the document.</summary>
    Top,
    /// <summary>Insert for the preceding heading.</summary>
    PreviousHeading
}

/// <summary>Adds a Markdown table of contents placeholder.</summary>
/// <example>
///   <summary>Add a TOC at the top of a generated report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\Report.md {
///     Add-OfficeMarkdownHeading -Level 1 -Text 'Release Report'
///     Add-OfficeMarkdownHeading -Level 2 -Text 'Summary'
///     Add-OfficeMarkdownParagraph -Text 'Current state and risks.'
///     Add-OfficeMarkdownHeading -Level 2 -Text 'Appendix'
///     Add-OfficeMarkdownParagraph -Text 'Supporting evidence.'
///     Add-OfficeMarkdownTableOfContents -Title 'Contents' -MinLevel 2 -MaxLevel 3 -Placement Top
/// }</code>
///   <para>Inserts a generated table of contents for headings in the document.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownTableOfContents", DefaultParameterSetName = "Placement")]
[OutputType(typeof(MarkdownDocument))]
[Alias("MarkdownTableOfContents") ]
public sealed class AddOfficeMarkdownTableOfContentsCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "DocumentPlacement")]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "DocumentSection")]
    public MarkdownDocument Document { get; set; } = null!;

    /// <summary>Heading text displayed above the generated table of contents.</summary>
    [Parameter]
    public string Title { get; set; } = "Contents";

    /// <summary>Minimum heading depth included in the table of contents.</summary>
    [Parameter]
    public int MinLevel { get; set; } = 1;

    /// <summary>Maximum heading depth included in the table of contents.</summary>
    [Parameter]
    public int MaxLevel { get; set; } = 3;

    /// <summary>Generate an ordered table of contents list.</summary>
    [Parameter]
    public SwitchParameter Numbered { get; set; }

    /// <summary>Heading level used for the TOC title.</summary>
    [Parameter]
    public int TitleLevel { get; set; } = 2;

    /// <summary>Where to place the TOC.</summary>
    [Parameter(ParameterSetName = "Placement")]
    [Parameter(ParameterSetName = "DocumentPlacement")]
    public MarkdownTocPlacement Placement { get; set; }

    /// <summary>Scope the TOC to the named section heading.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Section")]
    [Parameter(Mandatory = true, ParameterSetName = "DocumentSection")]
    public string? Section { get; set; }

    /// <summary>Emit the updated Markdown document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        ValidateParameters();

        if (!string.IsNullOrWhiteSpace(Section))
        {
            doc.TocForSection(Section!, Title, MinLevel, MaxLevel, Numbered.IsPresent, TitleLevel);
        }
        else if (Placement == MarkdownTocPlacement.PreviousHeading)
        {
            doc.TocForPreviousHeading(Title, MinLevel, MaxLevel, Numbered.IsPresent, TitleLevel);
        }
        else if (Placement == MarkdownTocPlacement.Top)
        {
            doc.TocAtTop(Title, MinLevel, MaxLevel, Numbered.IsPresent, TitleLevel);
        }
        else
        {
            doc.Toc(opts =>
            {
                opts.Title = Title;
                opts.MinLevel = MinLevel;
                opts.MaxLevel = MaxLevel;
                opts.Ordered = Numbered.IsPresent;
                opts.TitleLevel = TitleLevel;
            });
        }

        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
    }

    private void ValidateParameters()
    {
        if (MinLevel < 1 || MinLevel > 6)
        {
            throw new PSArgumentOutOfRangeException(nameof(MinLevel), MinLevel, "MinLevel must be between 1 and 6.");
        }

        if (MaxLevel < 1 || MaxLevel > 6)
        {
            throw new PSArgumentOutOfRangeException(nameof(MaxLevel), MaxLevel, "MaxLevel must be between 1 and 6.");
        }

        if (MaxLevel < MinLevel)
        {
            throw new PSArgumentException("MaxLevel must be greater than or equal to MinLevel.");
        }

        if (TitleLevel < 1 || TitleLevel > 6)
        {
            throw new PSArgumentOutOfRangeException(nameof(TitleLevel), TitleLevel, "TitleLevel must be between 1 and 6.");
        }

    }

    private MarkdownDocument ResolveDocument()
    {
        if (ParameterSetName.StartsWith("Document", StringComparison.Ordinal))
        {
            return Document ?? throw new PSArgumentException("Provide a Markdown document.");
        }

        return MarkdownDslContext.Require(this).Document;
    }
}
