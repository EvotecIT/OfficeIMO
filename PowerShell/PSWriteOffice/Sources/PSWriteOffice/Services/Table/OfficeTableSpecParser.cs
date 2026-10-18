using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Services.Table;

internal static class OfficeTableSpecParser
{
    private static readonly string[] TextKeys = { "Text", "Value", "Content" };
    private static readonly string[] RunKeys = { "Run", "Runs" };
    private static readonly string[] ColumnSpanKeys = { "ColumnSpan", "ColSpan" };
    private static readonly string[] RowSpanKeys = { "RowSpan" };
    private static readonly string[] TextColorKeys = { "TextColor", "Color", "FontColor" };
    private static readonly string[] FillColorKeys = { "FillColor", "BackgroundColor", "CellFill" };
    private static readonly string[] FontSizeKeys = { "FontSize" };
    private static readonly string[] BoldKeys = { "Bold" };
    private static readonly string[] ItalicKeys = { "Italic" };
    private static readonly string[] UnderlineKeys = { "Underline", "Underlined" };
    private static readonly string[] UnderlineStyleKeys = { "UnderlineStyle", "UnderlineKind" };
    private static readonly string[] StrikeKeys = { "Strike", "Strikethrough" };
    private static readonly string[] AlignKeys = { "Align", "Alignment", "HorizontalAlign" };
    private static readonly string[] VerticalAlignKeys = { "VerticalAlign", "VerticalAlignment" };
    private static readonly string[] CellKeys = TextKeys
        .Concat(RunKeys)
        .Concat(ColumnSpanKeys)
        .Concat(RowSpanKeys)
        .Concat(TextColorKeys)
        .Concat(FillColorKeys)
        .Concat(FontSizeKeys)
        .Concat(BoldKeys)
        .Concat(ItalicKeys)
        .Concat(UnderlineKeys)
        .Concat(UnderlineStyleKeys)
        .Concat(StrikeKeys)
        .Concat(AlignKeys)
        .Concat(VerticalAlignKeys)
        .ToArray();

    public static bool TryCreate(
        IReadOnlyList<object> rows,
        string[]? propertyNames,
        string[]? header,
        out OfficeTableSpec spec,
        PowerShellObjectNormalizerOptions? normalizationOptions = null)
    {
        spec = null!;
        normalizationOptions ??= PowerShellObjectNormalizerOptions.Default;
        if (rows.Count == 0)
        {
            return false;
        }

        var hasExplicitRows = rows.All(IsExplicitRow);
        var hasStructuredMarker = rows.Any(ContainsStructuredCellMarker);
        if (!hasExplicitRows && !hasStructuredMarker)
        {
            return false;
        }

        var tableRows = new List<IReadOnlyList<OfficeTableCellSpec>>();
        int? headerRowIndex = null;
        var allowDefaultHeader = header == null;
        if (header is { Length: > 0 })
        {
            tableRows.Add(header.Select(static value => new OfficeTableCellSpec(value)).ToArray());
            headerRowIndex = 0;
        }

        string[]? columns = propertyNames;
        foreach (var row in rows)
        {
            if (TryCreateCell(row, requireExplicitCellShape: true, allowStyleOnlyCell: false, allowRunEnumerable: false, out var singleCell))
            {
                tableRows.Add(new[] { singleCell });
                continue;
            }

            if (IsExplicitRow(row))
            {
                tableRows.Add(CreateExplicitRow(row, normalizationOptions));
                continue;
            }

            if (PowerShellObjectNormalizer.TryProjectItem(row, columns, out var projectedColumns, out var values, normalizationOptions))
            {
                columns ??= projectedColumns;
                if (allowDefaultHeader && !headerRowIndex.HasValue && columns.Length > 0)
                {
                    headerRowIndex = 0;
                    tableRows.Insert(0, columns.Select(static value => new OfficeTableCellSpec(value)).ToArray());
                }

                tableRows.Add(values.Select(value => ToCell(value, normalizationOptions)).ToArray());
                continue;
            }

            tableRows.Add(new[] { ToCell(row, normalizationOptions) });
        }

        if (tableRows.Count == 0)
        {
            return false;
        }

        spec = new OfficeTableSpec(tableRows, headerRowIndex);
        return spec.RowCount > 0 && spec.ColumnCount > 0;
    }

    internal static bool TryCreateCell(object? value, out OfficeTableCellSpec spec)
        => TryCreateCell(value, requireExplicitCellShape: false, allowStyleOnlyCell: true, allowRunEnumerable: true, out spec);

    private static OfficeTableCellSpec[] CreateExplicitRow(object row, PowerShellObjectNormalizerOptions normalizationOptions)
    {
        var cells = new List<OfficeTableCellSpec>();
        foreach (var value in Enumerate(row))
        {
            cells.Add(TryCreateCell(value, requireExplicitCellShape: false, allowStyleOnlyCell: true, allowRunEnumerable: true, out var spec)
                ? spec
                : ToCell(value, normalizationOptions));
        }

        return cells.ToArray();
    }

    private static bool ContainsStructuredCellMarker(object row)
    {
        if (TryCreateCell(row, requireExplicitCellShape: true, allowStyleOnlyCell: false, allowRunEnumerable: false, out var cell) && cell.HasStructuredMarker)
        {
            return true;
        }

        if (!IsExplicitRow(row))
        {
            return false;
        }

        foreach (var value in Enumerate(row))
        {
            if (TryCreateCell(value, requireExplicitCellShape: false, allowStyleOnlyCell: true, allowRunEnumerable: true, out cell) && cell.HasStructuredMarker)
            {
                return true;
            }
        }

        return false;
    }

    private static OfficeTableCellSpec ToCell(object? value, PowerShellObjectNormalizerOptions normalizationOptions)
        => new(PowerShellObjectNormalizer.NormalizeCellText(UnwrapPSObject(value), normalizationOptions));

    private static bool TryCreateCell(
        object? value,
        bool requireExplicitCellShape,
        bool allowStyleOnlyCell,
        bool allowRunEnumerable,
        out OfficeTableCellSpec spec)
    {
        value = UnwrapPSObject(value);
        if (value is OfficeTableCellSpec typed)
        {
            spec = typed;
            return true;
        }

        var hasColumnSpan = TryGetValue(value, ColumnSpanKeys, out var columnSpanValue);
        var hasRowSpan = TryGetValue(value, RowSpanKeys, out var rowSpanValue);
        var hasSpan = hasColumnSpan || hasRowSpan;
        var hasText = TryGetValue(value, TextKeys, out var textValue);
        var hasRuns = TryGetValue(value, RunKeys, out var runValue) && IsRunValue(runValue, allowStringRun: !requireExplicitCellShape, allowEnumerableRun: allowRunEnumerable);
        var hasOnlyCellKeys = HasOnlyCellKeys(value);
        if (!hasOnlyCellKeys)
        {
            spec = null!;
            return false;
        }

        if (!allowStyleOnlyCell && !hasSpan && !hasRuns)
        {
            spec = null!;
            return false;
        }

        var style = CreateStyle(value);
        var hasStyle = style?.HasAnyValue == true;
        if (!allowStyleOnlyCell && !hasSpan && hasStyle && !hasRuns)
        {
            spec = null!;
            return false;
        }

        if (requireExplicitCellShape && !hasSpan && !hasStyle && !hasRuns)
        {
            spec = null!;
            return false;
        }

        if (!hasSpan && !hasStyle && !hasText && !hasRuns)
        {
            spec = null!;
            return false;
        }

        if ((hasSpan || hasStyle) && !hasText && !hasRuns)
        {
            spec = null!;
            return false;
        }

        var runs = hasRuns ? OfficeTextRunParser.ParseMany(runValue) : null;
        var text = hasText
            ? Convert.ToString(UnwrapPSObject(textValue), CultureInfo.InvariantCulture)
            : runs != null
                ? OfficeTextRunParser.GetPlainText(runs)
            : string.Empty;
        var columnSpan = hasColumnSpan
            ? ConvertToPositiveInt(columnSpanValue, "ColumnSpan")
            : 1;
        var rowSpan = hasRowSpan
            ? ConvertToPositiveInt(rowSpanValue, "RowSpan")
            : 1;
        if (requireExplicitCellShape && columnSpan == 1 && rowSpan == 1 && !hasStyle && !hasRuns)
        {
            spec = null!;
            return false;
        }

        spec = new OfficeTableCellSpec(text, columnSpan, rowSpan, style, runs);
        return true;
    }

    private static OfficeTableCellStyle? CreateStyle(object? value)
    {
        var style = new OfficeTableCellStyle
        {
            TextColor = GetString(value, TextColorKeys),
            FillColor = GetString(value, FillColorKeys),
            FontSize = GetDouble(value, FontSizeKeys),
            Bold = GetBool(value, BoldKeys),
            Italic = GetBool(value, ItalicKeys),
            Underline = GetBool(value, UnderlineKeys) || GetString(value, UnderlineStyleKeys) != null,
            UnderlineStyle = GetString(value, UnderlineStyleKeys),
            Strike = GetBool(value, StrikeKeys),
            Align = GetString(value, AlignKeys),
            VerticalAlign = GetString(value, VerticalAlignKeys)
        };

        return style.HasAnyValue ? style : null;
    }

    private static bool HasOnlyCellKeys(object? source)
    {
        source = UnwrapPSObject(source);
        if (source is IDictionary dictionary)
        {
            return dictionary.Count > 0 &&
                   dictionary.Keys
                       .Cast<object?>()
                       .Select(static key => Convert.ToString(key, CultureInfo.InvariantCulture))
                       .All(IsCellKey);
        }

        if (source is null)
        {
            return false;
        }

        var psObject = PSObject.AsPSObject(source);
        if (psObject.BaseObject is IDictionary baseDictionary)
        {
            return HasOnlyCellKeys(baseDictionary);
        }

        var properties = psObject.Properties
            .Where(static property =>
                property.IsGettable &&
                property.MemberType is PSMemberTypes.NoteProperty or PSMemberTypes.Property or PSMemberTypes.AliasProperty)
            .ToArray();
        return properties.Length > 0 && properties.All(static property => IsCellKey(property.Name));
    }

    private static bool IsRunValue(object? value, bool allowStringRun, bool allowEnumerableRun)
    {
        value = UnwrapPSObject(value);
        if (value is null)
        {
            return false;
        }

        if (value is string)
        {
            return allowStringRun;
        }

        if (value is OfficeTextRunSpec or IDictionary)
        {
            return true;
        }

        return allowEnumerableRun &&
               value is IEnumerable enumerable &&
               enumerable.Cast<object?>().Any(static item => item is not null);
    }

    private static bool IsCellKey(string? key)
        => !string.IsNullOrWhiteSpace(key) &&
           CellKeys.Any(cellKey => string.Equals(cellKey, key, StringComparison.OrdinalIgnoreCase));

    private static bool IsExplicitRow(object row)
    {
        row = UnwrapPSObject(row)!;
        return row is IEnumerable and not string and not IDictionary and not DataTable and not DataView and not IDataReader and not DataSet &&
               !PowerShellDictionaryAdapter.IsDictionaryLike(row);
    }

    private static IEnumerable<object?> Enumerate(object row)
    {
        foreach (var value in (IEnumerable)UnwrapPSObject(row)!)
        {
            yield return value;
        }
    }

    private static bool TryGetValue(
        object? source,
        IEnumerable<string> keys,
        out object? value)
    {
        if (source is IDictionary dictionary)
        {
            foreach (var key in keys)
            {
                if (dictionary.Contains(key))
                {
                    value = dictionary[key];
                    return true;
                }

                foreach (DictionaryEntry entry in dictionary)
                {
                    if (string.Equals(Convert.ToString(entry.Key, CultureInfo.InvariantCulture), key, StringComparison.OrdinalIgnoreCase))
                    {
                        value = entry.Value;
                        return true;
                    }
                }
            }

            value = null;
            return false;
        }

        if (source is null)
        {
            value = null;
            return false;
        }

        var psObject = PSObject.AsPSObject(source);
        if (psObject.BaseObject is IDictionary baseDictionary)
        {
            return TryGetValue(baseDictionary, keys, out value);
        }

        foreach (var key in keys)
        {
            var property = psObject.Properties[key];
            if (property == null ||
                !property.IsGettable ||
                property.MemberType is not (PSMemberTypes.NoteProperty or PSMemberTypes.Property or PSMemberTypes.AliasProperty))
            {
                continue;
            }

            try
            {
                value = property.Value;
                return true;
            }
            catch (Exception exception) when (exception is not PipelineStoppedException)
            {
                value = null;
                return false;
            }
        }

        value = null;
        return false;
    }

    private static int ConvertToPositiveInt(object? value, string name)
    {
        value = UnwrapPSObject(value);
        var result = (int)LanguagePrimitives.ConvertTo(value, typeof(int), CultureInfo.InvariantCulture);
        if (result < 1)
        {
            throw new ArgumentOutOfRangeException(name, "Span values must be at least 1.");
        }

        return result;
    }

    private static string? GetString(object? source, IEnumerable<string> keys)
        => TryGetValue(source, keys, out var value)
            ? Convert.ToString(UnwrapPSObject(value), CultureInfo.InvariantCulture)
            : null;

    private static double? GetDouble(object? source, IEnumerable<string> keys)
        => TryGetValue(source, keys, out var value)
            ? Convert.ToDouble(UnwrapPSObject(value), CultureInfo.InvariantCulture)
            : null;

    private static bool GetBool(object? source, IEnumerable<string> keys)
        => TryGetValue(source, keys, out var value) &&
           Convert.ToBoolean(UnwrapPSObject(value), CultureInfo.InvariantCulture);

    private static object? UnwrapPSObject(object? value)
        => value is PSObject psObject ? psObject.BaseObject : value;

}
