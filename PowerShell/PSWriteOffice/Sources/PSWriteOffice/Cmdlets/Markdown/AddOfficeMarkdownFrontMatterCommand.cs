using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Markdown;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Adds YAML front matter to a Markdown document.</summary>
/// <example>
///   <summary>Add front matter from a hashtable.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>MarkdownFrontMatter -InputObject @{ title = 'Weekly Report'; tags = @('ops','summary') }</code>
///   <para>Sets the document header using the supplied key/value pairs.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeMarkdownFrontMatter", DefaultParameterSetName = ParameterSetContext)]
[Alias("MarkdownFrontMatter")]
[OutputType(typeof(MarkdownDoc))]
public sealed class AddOfficeMarkdownFrontMatterCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Front matter data expressed as a hashtable or object.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("Data")]
    public object InputObject { get; set; } = null!;

    /// <summary>Emit the updated Markdown document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var doc = ResolveDocument();
        doc.FrontMatter(NormalizeInputObject(InputObject));

        if (PassThru.IsPresent)
        {
            WriteObject(doc);
        }
    }

    private static object NormalizeInputObject(object? value)
    {
        if (value == null)
        {
            throw new PSArgumentNullException(nameof(InputObject));
        }

        if (value is IDictionary dictionary)
        {
            return NormalizeDictionary(dictionary);
        }

        var psObject = PSObject.AsPSObject(value);
        var properties = psObject.Properties
            .Where(p => p.MemberType == PSMemberTypes.NoteProperty || p.MemberType == PSMemberTypes.Property)
            .Where(p => !string.IsNullOrWhiteSpace(p.Name))
            .ToList();

        if (properties.Count == 0)
        {
            return value;
        }

        var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in properties)
        {
            result[property.Name] = NormalizeNestedValue(property.Value);
        }

        return result;
    }

    private static Dictionary<string, object?> NormalizeDictionary(IDictionary dictionary)
    {
        var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (DictionaryEntry entry in dictionary)
        {
            var key = entry.Key?.ToString();
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }

            result[key!] = NormalizeNestedValue(entry.Value);
        }

        return result;
    }

    private static object? NormalizeNestedValue(object? value)
    {
        if (value == null)
        {
            return null;
        }

        if (value is string or bool or byte or sbyte or short or ushort or int or uint or long or ulong
            or float or double or decimal or DateTime or DateTimeOffset or Guid)
        {
            return value;
        }

        if (value is IDictionary dictionary)
        {
            return NormalizeDictionary(dictionary);
        }

        if (value is IEnumerable enumerable && value is not string)
        {
            var items = new List<object?>();
            foreach (var item in enumerable)
            {
                items.Add(NormalizeNestedValue(item));
            }

            return items;
        }

        return value;
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
