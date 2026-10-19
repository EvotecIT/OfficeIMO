using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Appends or inserts a row in an existing PowerPoint table.</summary>
/// <para>
/// Accepts a <see cref="PowerPointTable"/> or a <see cref="PowerPointShapeInfo"/> record whose shape is a
/// table. Values can be scalars, arrays, dictionaries, or PowerShell objects; arrays and object
/// properties are expanded across table cells. The new row is cloned from an existing template row so
/// table formatting, borders, and style choices are preserved.
/// </para>
/// <example>
///   <summary>Find a table and append a row.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Find-OfficePowerPointShape -Presentation $ppt -Text 'Metric' -Kind Table |
///     Add-OfficePowerPointTableRow -Values 'Latency', 'Ready'</code>
///   <para>Accepts a PowerPoint table or table shape metadata and writes the supplied values into the new row.</para>
/// </example>
/// <example>
///   <summary>Insert a row above the first data row using a template row.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$shape = Find-OfficePowerPointShape -Presentation $ppt -Text 'Metric' -Kind Table | Select-Object -First 1
/// $shape | Add-OfficePowerPointTableRow -Index 1 -TemplateRowIndex 1 -Values ([ordered]@{
///     Metric = 'Documentation'
///     State  = 'Ready'
/// })</code>
///   <para>Resolves the table from shape metadata, inserts a formatted row at index 1, and maps values across cells.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointTableRow")]
[OutputType(typeof(PowerPointTableRow))]
public sealed class AddOfficePowerPointTableRowCommand : PSCmdlet
{
    /// <summary>PowerPoint table or table shape info returned by <c>Find-OfficePowerPointShape</c> or <c>Get-OfficePowerPointShape</c>.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public object InputObject { get; set; } = null!;

    /// <summary>Values to write into the new row. Arrays, dictionaries, and objects are expanded across cells.</summary>
    [Parameter(Position = 1)]
    [Alias("Data", "Values")]
    [AllowNull]
    public object? Value { get; set; }

    /// <summary>Optional zero-based template row index to clone. Defaults to the last existing row.</summary>
    [Parameter]
    public int? TemplateRowIndex { get; set; }

    /// <summary>Optional zero-based index where the row should be inserted. Defaults to appending at the end.</summary>
    [Parameter]
    public int? Index { get; set; }

    /// <summary>Emit the created row for additional OfficeIMO-level edits.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var table = ResolveTable(InputObject);
        var values = ExpandValues(Value);
        if (table.Rows > 0)
        {
            ValidateValues(values, table.Columns);
        }

        PowerPointTableRow row;
        if (table.Rows == 0)
        {
            table.AddRow(Index);
            row = table.GetRow(0);
        }
        else
        {
            var templateRowIndex = TemplateRowIndex ?? table.Rows - 1;
            row = table.AddRowFromTemplate(templateRowIndex, Index, clearText: true);
        }

        var occupiedColumns = new bool[row.Cells.Count];
        var valueIndex = 0;
        for (var column = 0; column < row.Cells.Count; column++)
        {
            if (occupiedColumns[column])
            {
                continue;
            }

            var cell = row.GetCell(column);
            if (valueIndex < values.Count)
            {
                var columnSpan = ApplyValue(cell, values[valueIndex++], column, row.Cells.Count);
                for (var offset = 1; offset < columnSpan && column + offset < occupiedColumns.Length; offset++)
                {
                    occupiedColumns[column + offset] = true;
                }
            }
            else
            {
                cell.Text = string.Empty;
            }
        }

        if (PassThru.IsPresent)
        {
            WriteObject(row);
        }
    }

    private static PowerPointTable ResolveTable(object input)
    {
        if (input is PSObject psObject)
        {
            input = psObject.BaseObject;
        }

        return input switch
        {
            PowerPointTable table => table,
            PowerPointShapeInfo { Shape: PowerPointTable table } => table,
            _ => throw new PSArgumentException("Input object must be a PowerPoint table or shape info for a table.", nameof(InputObject))
        };
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
        return value == null ? string.Empty : LanguagePrimitives.ConvertTo<string>(value) ?? string.Empty;
    }

    private static void ValidateValues(IReadOnlyList<object?> values, int columnCount)
    {
        var occupiedColumns = new bool[columnCount];
        var valueIndex = 0;
        for (var column = 0; column < columnCount && valueIndex < values.Count; column++)
        {
            if (occupiedColumns[column])
            {
                continue;
            }

            if (values[valueIndex] != null && OfficeTableSpecParser.TryCreateCell(values[valueIndex], out var spec))
            {
                ValidateSpan(spec, column, columnCount);
                PowerPointTableCellSpecService.Validate(spec);
                for (var offset = 1; offset < spec.ColumnSpan && column + offset < occupiedColumns.Length; offset++)
                {
                    occupiedColumns[column + offset] = true;
                }
            }

            valueIndex++;
        }
    }

    private static int ApplyValue(PowerPointTableCell cell, object? value, int column, int columnCount)
    {
        if (value != null && OfficeTableSpecParser.TryCreateCell(value, out var spec))
        {
            ValidateSpan(spec, column, columnCount);
            PowerPointTableCellSpecService.Apply(cell, spec);
            if (spec.HasSpan)
            {
                cell.Merge = (spec.RowSpan, spec.ColumnSpan);
            }

            return spec.ColumnSpan;
        }

        cell.Text = ConvertValue(value);
        return 1;
    }

    private static void ValidateSpan(OfficeTableCellSpec spec, int column, int columnCount)
    {
        if (spec.RowSpan > 1)
        {
            throw new PSArgumentException("Add-OfficePowerPointTableRow cannot apply RowSpan values because it creates one physical row.", nameof(Value));
        }

        if (column + spec.ColumnSpan > columnCount)
        {
            throw new PSArgumentException($"ColumnSpan {spec.ColumnSpan} at column {column} exceeds the table width of {columnCount}.", nameof(Value));
        }
    }
}
