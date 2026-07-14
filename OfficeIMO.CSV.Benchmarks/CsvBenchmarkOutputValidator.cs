#nullable enable

namespace OfficeIMO.CSV.Benchmarks;

internal static class CsvBenchmarkOutputValidator
{
    internal static void Validate(
        string method,
        string csvText,
        IReadOnlyList<string> expectedHeaders,
        int expectedRowCount,
        string?[][]? expectedRows)
    {
        if (csvText.Length == 0)
        {
            throw new InvalidOperationException($"{method} produced an empty CSV document.");
        }

        var rowIndex = 0;
        using var reader = new StringReader(csvText);
        CsvDocument.ReadRowsReusable(reader, (headers, values) =>
        {
            if (rowIndex == 0 && !headers.SequenceEqual(expectedHeaders, StringComparer.Ordinal))
            {
                throw new InvalidOperationException($"{method} produced a different CSV header.");
            }

            if (values.Count != expectedHeaders.Count)
            {
                throw new InvalidOperationException(
                    $"{method} row {rowIndex + 1} contains {values.Count} fields; expected {expectedHeaders.Count}.");
            }

            if (expectedRows != null)
            {
                if (rowIndex >= expectedRows.Length)
                {
                    throw new InvalidOperationException($"{method} produced more than {expectedRows.Length} data rows.");
                }

                var expected = expectedRows[rowIndex];
                for (var fieldIndex = 0; fieldIndex < expected.Length; fieldIndex++)
                {
                    var expectedValue = expected[fieldIndex] ?? string.Empty;
                    if (!string.Equals(values[fieldIndex], expectedValue, StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            $"{method} row {rowIndex + 1}, field {fieldIndex + 1} did not round-trip the prepared text value.");
                    }
                }
            }

            rowIndex++;
        });

        if (rowIndex != expectedRowCount)
        {
            throw new InvalidOperationException($"{method} produced {rowIndex} data rows; expected {expectedRowCount}.");
        }
    }
}
