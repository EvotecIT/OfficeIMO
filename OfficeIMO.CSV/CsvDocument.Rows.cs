#nullable enable

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    private object?[] NormalizeValuesLength(IEnumerable<object?> values)
    {
        var array = values.ToArray();
        if (_header.Count == 0)
        {
            _header.AddRange(GenerateDefaultHeader(array.Length));
            return array;
        }

        if (array.Length != _header.Count)
        {
            throw new CsvException($"Row contains {array.Length} values but header defines {_header.Count} columns.");
        }

        return array;
    }

    private IEnumerable<object?[]> EnumerateRawRows()
    {
        if (_mode == CsvLoadMode.Stream && _streamingSource is not null)
        {
            return _streamingSource.ReadRows();
        }

        return EnumerateInMemoryRows();

        IEnumerable<object?[]> EnumerateInMemoryRows()
        {
            foreach (var row in _rows)
            {
                yield return row.Values;
            }
        }
    }

    private static IReadOnlyList<string> GenerateDefaultHeader(int count)
    {
        var result = new List<string>(count);
        for (var i = 0; i < count; i++)
        {
            result.Add($"Column{i + 1}");
        }

        return result;
    }

    private static IReadOnlyList<string> NormalizeParsedHeader(IReadOnlyList<string> header, CsvLoadOptions options)
    {
        if (!options.GenerateMissingHeaderNames && options.DuplicateHeaderBehavior == CsvDuplicateHeaderBehavior.Preserve)
        {
            return header.ToArray();
        }

        var result = new string[header.Count];
        var sourceNames = CreateHeaderNameSet(header);
        var assigned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var generated = 1;
        for (var i = 0; i < header.Count; i++)
        {
            var name = header[i];
            if (string.IsNullOrEmpty(name) && options.GenerateMissingHeaderNames)
            {
                do
                {
                    name = $"H{generated++}";
                }
                while (sourceNames.Contains(name) || assigned.Contains(name));
            }

            if (!string.IsNullOrEmpty(name) && !assigned.Add(name))
            {
                name = options.DuplicateHeaderBehavior switch
                {
                    CsvDuplicateHeaderBehavior.Preserve => name,
                    CsvDuplicateHeaderBehavior.Rename => CreateUniqueDuplicateHeaderName(name, sourceNames, assigned),
                    CsvDuplicateHeaderBehavior.Throw => throw new CsvException($"CSV header contains duplicate column name '{name}'."),
                    _ => throw new ArgumentOutOfRangeException(nameof(options), options.DuplicateHeaderBehavior, "Unsupported duplicate CSV header behavior.")
                };

                assigned.Add(name);
            }

            result[i] = name;
        }

        return result;
    }

    private static HashSet<string> CreateHeaderNameSet(IReadOnlyList<string> header)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < header.Count; i++)
        {
            if (!string.IsNullOrEmpty(header[i]))
            {
                names.Add(header[i]);
            }
        }

        return names;
    }

    private static string CreateUniqueDuplicateHeaderName(string name, HashSet<string> sourceHeader, HashSet<string> assignedHeader)
    {
        var suffix = 2;
        string candidate;
        do
        {
            candidate = $"{name}_{suffix++}";
        }
        while (sourceHeader.Contains(candidate) || assignedHeader.Contains(candidate));

        return candidate;
    }

    internal static IReadOnlyList<string> AppendStaticColumnsToHeader(IReadOnlyList<string> header, CsvLoadOptions options)
    {
        if (options.StaticColumns is null || options.StaticColumns.Count == 0)
        {
            return header;
        }

        var combined = new string[header.Count + options.StaticColumns.Count];
        for (var i = 0; i < header.Count; i++)
        {
            combined[i] = header[i];
        }

        var index = header.Count;
        foreach (var staticColumn in options.StaticColumns)
        {
            combined[index++] = staticColumn.Key;
        }

        return NormalizeParsedHeader(combined, options);
    }

    internal static IReadOnlyList<string> BuildParsedStringValues(IReadOnlyList<string> values, int headerCount, CsvLoadOptions options)
    {
        var staticCount = options.StaticColumns?.Count ?? 0;
        var sourceHeaderCount = headerCount - staticCount;
        var aligned = AlignParsedStringValues(values, sourceHeaderCount, options.ColumnCountMismatchPolicy);
        if (staticCount == 0)
        {
            return aligned;
        }

        var result = new string[headerCount];
        for (var i = 0; i < aligned.Count; i++)
        {
            result[i] = aligned[i];
        }

        var index = aligned.Count;
        foreach (var staticColumn in options.StaticColumns!)
        {
            result[index++] = Convert.ToString(staticColumn.Value, options.Culture) ?? string.Empty;
        }

        return result;
    }

    internal static object?[] BuildParsedObjectValues(IReadOnlyList<string> values, int headerCount, CsvLoadOptions options)
    {
        return FillParsedObjectValues(values, headerCount, options, target: null);
    }

    internal static object?[] FillParsedObjectValues(IReadOnlyList<string> values, int headerCount, CsvLoadOptions options, object?[]? target)
    {
        var staticCount = options.StaticColumns?.Count ?? 0;
        var sourceHeaderCount = headerCount - staticCount;
        var aligned = target is { Length: var length } && length == headerCount
            ? target
            : new object?[headerCount];

        var copyCount = Math.Min(values.Count, sourceHeaderCount);
        if (values.Count != sourceHeaderCount && options.ColumnCountMismatchPolicy == CsvColumnCountMismatchPolicy.Strict)
        {
            throw new CsvException($"Row contains {values.Count} values but header defines {sourceHeaderCount} columns.");
        }

        for (var i = 0; i < copyCount; i++)
        {
            aligned[i] = NormalizeLoadedValue(values[i], options);
        }

        for (var i = copyCount; i < sourceHeaderCount; i++)
        {
            aligned[i] = string.Empty;
        }

        if (staticCount > 0)
        {
            var index = sourceHeaderCount;
            foreach (var staticColumn in options.StaticColumns!)
            {
                aligned[index++] = staticColumn.Value;
            }
        }

        return aligned;
    }

    private static object? NormalizeLoadedValue(string value, CsvLoadOptions options)
    {
        return options.NullValue is not null && string.Equals(value, options.NullValue, StringComparison.Ordinal)
            ? null
            : value;
    }

    private static IReadOnlyList<string> AlignParsedStringValues(IReadOnlyList<string> values, int headerCount, CsvColumnCountMismatchPolicy policy)
    {
        if (values.Count == headerCount)
        {
            return values;
        }

        if (policy == CsvColumnCountMismatchPolicy.Strict)
        {
            throw new CsvException($"Row contains {values.Count} values but header defines {headerCount} columns.");
        }

        var aligned = new string[headerCount];
        var copyCount = Math.Min(values.Count, headerCount);
        for (var i = 0; i < copyCount; i++)
        {
            aligned[i] = values[i];
        }

        for (var i = copyCount; i < headerCount; i++)
        {
            aligned[i] = string.Empty;
        }

        return aligned;
    }

    private static object?[] AlignParsedObjectValues(IReadOnlyList<object?> values, int headerCount, CsvColumnCountMismatchPolicy policy)
    {
        if (values.Count == headerCount)
        {
            var exact = new object?[values.Count];
            for (var i = 0; i < values.Count; i++)
            {
                exact[i] = values[i];
            }

            return exact;
        }

        if (policy == CsvColumnCountMismatchPolicy.Strict)
        {
            throw new CsvException($"Row contains {values.Count} values but header defines {headerCount} columns.");
        }

        var aligned = new object?[headerCount];
        var copyCount = Math.Min(values.Count, headerCount);
        for (var i = 0; i < copyCount; i++)
        {
            aligned[i] = values[i];
        }

        for (var i = copyCount; i < headerCount; i++)
        {
            aligned[i] = string.Empty;
        }

        return aligned;
    }
}
