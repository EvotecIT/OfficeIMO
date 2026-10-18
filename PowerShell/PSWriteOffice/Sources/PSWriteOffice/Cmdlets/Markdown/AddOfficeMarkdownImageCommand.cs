using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown image.</summary>
/// <example>
///   <summary>Add a report image with alt text and sizing.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\Report.md {
///     Add-OfficeMarkdownHeading -Level 2 -Text 'Architecture'
///     Add-OfficeMarkdownImage -Path '.\Assets\architecture.png' -Alt 'Architecture diagram' -Title 'Current architecture' -Width 640
/// }</code>
///   <para>Appends an image block with accessibility text and optional presentation metadata.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownImage", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownImage")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownImageCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Image path or URL.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Alt text for the image.</summary>
    [Parameter]
    public string? Alt { get; set; }

    /// <summary>Optional title for the image.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Optional width in pixels.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Optional height in pixels.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Emit the Markdown document after appending the image.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        doc.Image(Path, Alt, Title, Width, Height);

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
