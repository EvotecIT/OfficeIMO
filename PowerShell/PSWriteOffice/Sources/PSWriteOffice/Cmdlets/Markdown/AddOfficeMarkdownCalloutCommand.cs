using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown callout block.</summary>
/// <example>
///   <summary>Add release callouts to a Markdown report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\ReleaseNotes.md {
///     Add-OfficeMarkdownHeading -Level 1 -Text 'Release notes'
///     Add-OfficeMarkdownCallout -Kind 'note' -Title 'Validation' -Body 'Artifacts were generated from deterministic example data.'
///     Add-OfficeMarkdownCallout -Kind 'warning' -Title 'Manual step' -Body 'Open the workbook in desktop Excel before publishing pivots.'
/// }</code>
///   <para>Appends callout blocks while composing a Markdown report.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownCallout", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownCallout")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownCalloutCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Callout kind (e.g. note, tip, warning).</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [ArgumentCompleter(typeof(OfficeMarkdownCalloutKindArgumentCompleter))]
    public string Kind { get; set; } = "note";

    /// <summary>Callout title.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Title { get; set; } = string.Empty;

    /// <summary>Callout body text.</summary>
    [Parameter(Mandatory = true, Position = 2)]
    public string Body { get; set; } = string.Empty;

    /// <summary>Emit the Markdown document after appending the callout.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        doc.Callout(Kind, Title, Body);

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
