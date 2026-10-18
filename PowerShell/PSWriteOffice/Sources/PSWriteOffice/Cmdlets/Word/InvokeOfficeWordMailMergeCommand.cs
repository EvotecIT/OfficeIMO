using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Executes a simple mail merge against MERGEFIELD values in a Word document.</summary>
/// <example>
///   <summary>Replace merge fields from a hashtable.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Invoke-OfficeWordMailMerge -InputObject @{ FirstName = 'John'; OrderId = 12345 }</code>
///   <para>Updates MERGEFIELD values in the active Word document.</para>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeWordMailMerge")]
[OutputType(typeof(WordDocument))]
public sealed class InvokeOfficeWordMailMergeCommand : PSCmdlet
{
    /// <summary>Document to update when provided explicitly.</summary>
    [Parameter(ValueFromPipeline = true)]
    public WordDocument? Document { get; set; }

    /// <summary>Hashtable or object whose properties map to MERGEFIELD names.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("Data", "Values")]
    public object InputObject { get; set; } = null!;

    /// <summary>Preserve field codes and only update displayed field text.</summary>
    [Parameter]
    public SwitchParameter PreserveFields { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = Document ?? WordDslContext.Require(this).Document;
        if (document == null)
        {
            throw new InvalidOperationException("Word document was not provided.");
        }

        var values = NormalizeInputObject(InputObject);
        if (values.Count == 0)
        {
            throw new PSArgumentException("Provide at least one mail-merge value.", nameof(InputObject));
        }

        WordMailMerge.Execute(document, values, removeFields: !PreserveFields.IsPresent);

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }

    private static Dictionary<string, string> NormalizeInputObject(object? value)
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
            .Where(property => property.MemberType == PSMemberTypes.NoteProperty || property.MemberType == PSMemberTypes.Property)
            .Where(property => !string.IsNullOrWhiteSpace(property.Name))
            .ToList();

        if (properties.Count == 0)
        {
            throw new PSArgumentException("Mail-merge input must be a hashtable or an object with readable properties.", nameof(InputObject));
        }

        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in properties)
        {
            result[property.Name] = ConvertToString(property.Value);
        }

        return result;
    }

    private static Dictionary<string, string> NormalizeDictionary(IDictionary dictionary)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (DictionaryEntry entry in dictionary)
        {
            var key = entry.Key?.ToString();
            if (string.IsNullOrWhiteSpace(key))
            {
                continue;
            }

            result[key!] = ConvertToString(entry.Value);
        }

        return result;
    }

    private static string ConvertToString(object? value)
    {
        if (value == null)
        {
            return string.Empty;
        }

        value = value is PSObject psObject ? psObject.BaseObject : value;
        if (value is DateTime dateTime)
        {
            return dateTime.ToString("o", CultureInfo.InvariantCulture);
        }

        if (value is DateTimeOffset dateTimeOffset)
        {
            return dateTimeOffset.ToString("o", CultureInfo.InvariantCulture);
        }

        return System.Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    }
}
