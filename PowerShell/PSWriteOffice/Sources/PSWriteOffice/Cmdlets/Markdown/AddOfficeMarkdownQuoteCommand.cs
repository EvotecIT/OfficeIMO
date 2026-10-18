using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown block quote.</summary>
/// <example>
///   <summary>Add an executive takeaway quote.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\ExecutiveSummary.md {
///     Add-OfficeMarkdownHeading -Level 1 -Text 'Executive summary'
///     Add-OfficeMarkdownQuote -Text 'The release is ready when examples, generated docs, and artifact readback all agree.'
///     Add-OfficeMarkdownParagraph -Text 'Detailed validation follows.'
/// }</code>
///   <para>Appends a quote block to the current Markdown document.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownQuote", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownQuote")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownQuoteCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Quote text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Emit the Markdown document after appending the quote.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        doc.Quote(Text ?? string.Empty);

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
