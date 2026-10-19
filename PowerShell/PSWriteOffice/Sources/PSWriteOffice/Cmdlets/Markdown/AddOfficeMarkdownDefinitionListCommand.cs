using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds a Markdown definition list.</summary>
/// <example>
///   <summary>Add term/definition pairs.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>MarkdownDefinitionList -Definition @{ SLA = 'Service level agreement'; SLO = 'Service level objective' }</code>
///   <para>Appends a definition list built from the provided pairs.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownDefinitionList", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownDefinitionList")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownDefinitionListCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Hashtable of term/definition pairs to render.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public Hashtable Definition { get; set; } = new();

    /// <summary>Emit the updated Markdown document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        var pairs = NormalizePairs(Definition);
        if (pairs.Count == 0)
        {
            throw new PSArgumentException("Provide at least one term/definition pair.", nameof(Definition));
        }

        doc.Dl(builder =>
        {
            foreach (var pair in pairs)
            {
                builder.Item(pair.Key, pair.Value);
            }
        });

        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
    }

    private static List<KeyValuePair<string, string>> NormalizePairs(Hashtable? definition)
    {
        var pairs = new List<KeyValuePair<string, string>>();
        if (definition == null)
        {
            return pairs;
        }

        foreach (DictionaryEntry entry in definition)
        {
            var term = entry.Key?.ToString()?.Trim();
            var body = entry.Value?.ToString()?.Trim();
            if (string.IsNullOrWhiteSpace(term) || string.IsNullOrWhiteSpace(body))
            {
                continue;
            }

            pairs.Add(new KeyValuePair<string, string>(term!, body!));
        }

        return pairs;
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
