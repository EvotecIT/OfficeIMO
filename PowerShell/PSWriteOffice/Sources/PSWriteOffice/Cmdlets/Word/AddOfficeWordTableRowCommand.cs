using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Appends a row to an existing Word table.</summary>
/// <para>
/// Adds a new row to a <see cref="WordTable"/> that was already created or found in an existing
/// document. The command accepts scalar values, arrays, dictionaries, ordered dictionaries, and
/// PowerShell objects. Values are expanded from left to right across cells; missing values become empty
/// cells. This keeps existing-document editing simple without forcing callers back into the Word DSL.
/// </para>
/// <example>
///   <summary>Append values to the first table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $doc | Get-OfficeWordTable | Select-Object -First 1 |
///     Add-OfficeWordTableRow -Values 'Service', 'Ready', 'Low'</code>
///   <para>Adds one table row and writes the supplied values into its cells.</para>
/// </example>
/// <example>
///   <summary>Append an object-like row to a table found by marker text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc = Get-OfficeWord -Path .\Report.docx
/// $table = Find-OfficeWordTable -Document $doc -Text 'Risk marker' | Select-Object -First 1
/// $table | Add-OfficeWordTableRow -Values ([ordered]@{
///     Item  = 'Mitigation plan'
///     Owner = 'Service Desk'
///     State = 'Ready'
/// })
/// $doc | Close-OfficeWord -Save</code>
///   <para>Uses an ordered dictionary so values are written into predictable table columns.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordTableRow")]
[OutputType(typeof(WordTableRow))]
public sealed class AddOfficeWordTableRowCommand : PSCmdlet
{
    /// <summary>Existing Word table to append to, usually from <c>Get-OfficeWordTable</c> or <c>Find-OfficeWordTable</c>.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public WordTable Table { get; set; } = null!;

    /// <summary>Values to write into the new row. Arrays, dictionaries, ordered dictionaries, and objects are expanded across cells.</summary>
    [Parameter(Position = 1)]
    [Alias("Data", "InputObject")]
    [AllowNull]
    public object? Values { get; set; }

    /// <summary>Emit the created row for additional OfficeIMO-level edits.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Table == null)
        {
            return;
        }

        var values = ExpandValues(Values);
        var existingColumnCount = Table.RowsCount > 0
            ? Table.Rows.Max(row => row.CellsCount)
            : 0;
        var cellCount = Math.Max(existingColumnCount, values.Count);
        if (cellCount == 0)
        {
            cellCount = 1;
        }

        var row = Table.AddRow(cellCount);
        for (var index = 0; index < row.CellsCount; index++)
        {
            var text = index < values.Count
                ? ConvertValue(values[index])
                : string.Empty;
            row.Cells[index].AddParagraph(text, removeExistingParagraphs: true);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(row);
        }
    }

    private static IReadOnlyList<object?> ExpandValues(object? value)
    {
        if (value == null)
        {
            return Array.Empty<object?>();
        }

        if (value is PSObject psObject)
        {
            if (psObject.BaseObject is IDictionary dictionary)
            {
                return ExpandDictionary(dictionary);
            }

            if (psObject.BaseObject is not string && psObject.BaseObject is IEnumerable enumerable)
            {
                return ExpandEnumerable(enumerable);
            }

            var properties = psObject.Properties
                .Where(property => property.MemberType is PSMemberTypes.NoteProperty or PSMemberTypes.Property)
                .Select(property => property.Value)
                .ToArray();
            if (properties.Length > 0)
            {
                return properties;
            }

            return new object?[] { psObject.BaseObject };
        }

        if (value is IDictionary dict)
        {
            return ExpandDictionary(dict);
        }

        if (value is not string && value is IEnumerable list)
        {
            return ExpandEnumerable(list);
        }

        return new object?[] { value };
    }

    private static IReadOnlyList<object?> ExpandDictionary(IDictionary dictionary)
    {
        var values = new List<object?>(dictionary.Count);
        foreach (DictionaryEntry entry in dictionary)
        {
            values.Add(entry.Value);
        }
        return values;
    }

    private static IReadOnlyList<object?> ExpandEnumerable(IEnumerable enumerable)
    {
        var values = new List<object?>();
        foreach (var item in enumerable)
        {
            values.Add(item);
        }
        return values;
    }

    private static string ConvertValue(object? value)
    {
        if (value == null)
        {
            return string.Empty;
        }

        return LanguagePrimitives.ConvertTo<string>(value) ?? string.Empty;
    }
}
