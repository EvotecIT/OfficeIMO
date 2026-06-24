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
        if (!options.GenerateMissingHeaderNames)
        {
            return header.ToArray();
        }

        var result = new string[header.Count];
        var generated = 1;
        for (var i = 0; i < header.Count; i++)
        {
            var name = header[i];
            if (string.IsNullOrEmpty(name))
            {
                do
                {
                    name = $"H{generated++}";
                }
                while (header.Contains(name, StringComparer.OrdinalIgnoreCase) || result.Contains(name, StringComparer.OrdinalIgnoreCase));
            }

            result[i] = name;
        }

        return result;
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
