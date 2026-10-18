using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a collapsible Markdown details block.</summary>
/// <example>
///   <summary>Add a details section.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>MarkdownDetails -Summary 'Implementation notes' { MarkdownParagraph -Text 'Hidden by default.' }</code>
///   <para>Appends a details/summary block with nested Markdown content.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownDetails", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownDetails")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownDetailsCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Summary text displayed by the details block.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Summary { get; set; } = string.Empty;

    /// <summary>Nested Markdown content rendered inside the details block.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public ScriptBlock Content { get; set; } = null!;

    /// <summary>Render the details block as open by default.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit the updated Markdown document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        doc.Details(Summary ?? string.Empty, inner =>
        {
            using (MarkdownDslContext.Enter(inner))
            {
                Content.InvokeReturnAsIs();
            }
        }, Open.IsPresent);

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
