#nullable enable

using System.Globalization;

namespace OfficeIMO.CSV.Benchmarks;

internal static class CsvBenchmarkOutputValidator
{
    internal static void Validate(
        string method,
        string csvText,
        IReadOnlyList<string> expectedHeaders,
        int expectedRowCount,
        string?[][]? expectedTextRows,
        object?[][]? expectedObjectRows = null)
    {
        if (expectedTextRows != null && expectedObjectRows != null)
        {
            throw new ArgumentException("Specify either prepared text rows or typed object rows, not both.");
        }

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

            if (expectedTextRows != null)
            {
                if (rowIndex >= expectedTextRows.Length)
                {
                    throw new InvalidOperationException($"{method} produced more than {expectedTextRows.Length} data rows.");
                }

                var expected = expectedTextRows[rowIndex];
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
            else if (expectedObjectRows != null)
            {
                if (rowIndex >= expectedObjectRows.Length)
                {
                    throw new InvalidOperationException($"{method} produced more than {expectedObjectRows.Length} data rows.");
                }

                var expected = expectedObjectRows[rowIndex];
                for (var fieldIndex = 0; fieldIndex < expected.Length; fieldIndex++)
                {
                    if (!SemanticallyEquals(values[fieldIndex], expected[fieldIndex]))
                    {
                        throw new InvalidOperationException(
                            $"{method} row {rowIndex + 1}, field {fieldIndex + 1} did not round-trip the typed value '{expected[fieldIndex]}'.");
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

    private static bool SemanticallyEquals(string actual, object? expected)
    {
        if (expected == null || ReferenceEquals(expected, DBNull.Value))
        {
            return actual.Length == 0;
        }

        switch (expected)
        {
            case string text:
                return string.Equals(actual, text, StringComparison.Ordinal);
            case bool boolean:
                return (bool.TryParse(actual, out var actualBoolean) && actualBoolean == boolean)
                    || (int.TryParse(actual, NumberStyles.Integer, CultureInfo.InvariantCulture, out var actualBooleanNumber)
                        && actualBooleanNumber == (boolean ? 1 : 0));
            case DateTime dateTime:
                return DateTime.TryParse(
                        actual,
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind,
                        out var actualDateTime)
                    && actualDateTime.Ticks == dateTime.Ticks;
            case DateTimeOffset dateTimeOffset:
                return DateTimeOffset.TryParse(
                        actual,
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind,
                        out var actualDateTimeOffset)
                    && actualDateTimeOffset.Equals(dateTimeOffset);
            case byte or sbyte or short or ushort or int or uint or long or ulong:
                return decimal.TryParse(actual, NumberStyles.Integer, CultureInfo.InvariantCulture, out var actualInteger)
                    && actualInteger == Convert.ToDecimal(expected, CultureInfo.InvariantCulture);
            case decimal decimalValue:
                return decimal.TryParse(
                        actual,
                        NumberStyles.Number | NumberStyles.AllowExponent,
                        CultureInfo.InvariantCulture,
                        out var actualNumber)
                    && actualNumber == decimalValue;
            case float singleValue:
                return float.TryParse(actual, NumberStyles.Float, CultureInfo.InvariantCulture, out var actualSingle)
                    && actualSingle.Equals(singleValue);
            case double doubleValue:
                return double.TryParse(actual, NumberStyles.Float, CultureInfo.InvariantCulture, out var actualDouble)
                    && actualDouble.Equals(doubleValue);
            default:
                return string.Equals(actual, Convert.ToString(expected, CultureInfo.InvariantCulture), StringComparison.Ordinal);
        }
    }
}
