using System.Globalization;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.CSV.Benchmarks;

internal static class CsvBenchmarkOutputValidator
{
    private static readonly (string Prefix, string Library)[] LibraryPrefixes =
    [
        ("OfficeIMO", "OfficeIMO"),
        ("ExcelReaderNet", "ExcelReader.NET"),
        ("CsvHelper", "CsvHelper"),
        ("Sylvan", "Sylvan.Data.Csv"),
        ("Sep", "Sep"),
        ("Dataplat", "Dataplat.Dbatools.Csv"),
        ("LumenWorks", "LumenWorks"),
        ("RecordParser", "RecordParser")
    ];

    /// <summary>
    /// The declared library and case of a lane named "Library_Case" (or just "Library", whose case is Default),
    /// matching what Build/BenchmarkDeclarations.ps1 derives for multi-library comparisons.
    /// </summary>
    internal static (string Library, string Case) Identify(string method)
    {
        foreach (var (prefix, library) in LibraryPrefixes)
        {
            if (method.Equals(prefix, StringComparison.Ordinal))
            {
                return (library, "Default");
            }

            if (method.StartsWith(prefix + "_", StringComparison.Ordinal))
            {
                return (library, method.Substring(prefix.Length + 1));
            }
        }

        int separator = method.IndexOf('_');
        return separator > 0 ? (method.Substring(0, separator), method.Substring(separator + 1)) : (method, "Default");
    }

    /// <summary>
    /// Runs one lane outside the measurement; the lane validates its own output and throws on a mismatch. The
    /// result is recorded for the benchmark evidence under the lane's library and case.
    /// </summary>
    internal static T Record<T>(string method, string detail, Func<T> lane)
    {
        var (library, caseName) = Identify(method);
        return RecordCore(library, caseName, $"{method}: {detail}", lane);
    }

    /// <summary>Runs one self-validating lane whose library and case are given explicitly, and records the result.</summary>
    internal static T Record<T>(string library, string caseName, string detail, Func<T> lane) =>
        RecordCore(library, caseName, detail, lane);

    private static T RecordCore<T>(string library, string caseName, string detail, Func<T> lane)
    {
        T result;
        try
        {
            result = lane();
        }
        catch (Exception exception)
        {
            BenchmarkOutputValidationLog.Failed(library, $"{detail}: {exception.Message}", caseName);
            throw;
        }

        BenchmarkOutputValidationLog.Passed(library, detail, caseName);
        return result;
    }

    /// <summary>Records that <paramref name="method"/>'s output was already checked by the caller.</summary>
    internal static void RecordPassed(string method, string detail)
    {
        var (library, caseName) = Identify(method);
        BenchmarkOutputValidationLog.Passed(library, $"{method}: {detail}", caseName);
    }

    /// <summary>Records that <paramref name="method"/>'s output did not match.</summary>
    internal static void RecordFailed(string method, string detail)
    {
        var (library, caseName) = Identify(method);
        BenchmarkOutputValidationLog.Failed(library, $"{method}: {detail}", caseName);
    }

    internal static void Validate(
        string method,
        string csvText,
        IReadOnlyList<string> expectedHeaders,
        int expectedRowCount,
        string?[][]? expectedTextRows,
        object?[][]? expectedObjectRows = null,
        string delimiter = ",")
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
        using (var reader = new StringReader(csvText))
        {
            CsvDocument.ReadRowsReusable(reader, (headers, values) =>
            {
                CheckRow(method, "OfficeIMO.CSV", rowIndex, headers, values, expectedHeaders, expectedTextRows, expectedObjectRows);
                rowIndex++;
            }, new CsvReaderOptions { DelimiterText = delimiter });
        }

        if (rowIndex != expectedRowCount)
        {
            throw new InvalidOperationException($"{method} produced {rowIndex} data rows; expected {expectedRowCount}.");
        }

        // Read the same text again with an independent parser, so a writer is never checked only by the reader of
        // the library under test. Sylvan.Data.Csv accepts single-character delimiters.
        if (delimiter.Length == 1)
        {
            ValidateWithSylvan(method, csvText, expectedHeaders, expectedRowCount, expectedTextRows, expectedObjectRows, delimiter[0]);
        }
    }

    private static void ValidateWithSylvan(
        string method,
        string csvText,
        IReadOnlyList<string> expectedHeaders,
        int expectedRowCount,
        string?[][]? expectedTextRows,
        object?[][]? expectedObjectRows,
        char delimiter)
    {
        using var text = new StringReader(csvText);
        using var reader = Sylvan.Data.Csv.CsvDataReader.Create(text, new Sylvan.Data.Csv.CsvDataReaderOptions
        {
            HasHeaders = true,
            Delimiter = delimiter,
            BufferSize = 1 << 20
        });
        var headers = new string[reader.FieldCount];
        for (var ordinal = 0; ordinal < headers.Length; ordinal++)
        {
            headers[ordinal] = reader.GetName(ordinal);
        }

        var values = new List<string>(headers.Length);
        var rowIndex = 0;
        while (reader.Read())
        {
            values.Clear();
            for (var ordinal = 0; ordinal < reader.RowFieldCount; ordinal++)
            {
                values.Add(reader.GetString(ordinal));
            }

            CheckRow(method, "Sylvan.Data.Csv", rowIndex, headers, values, expectedHeaders, expectedTextRows, expectedObjectRows);
            rowIndex++;
        }

        if (rowIndex != expectedRowCount)
        {
            throw new InvalidOperationException($"{method} produced {rowIndex} data rows when read by Sylvan.Data.Csv; expected {expectedRowCount}.");
        }
    }

    private static void CheckRow(
        string method,
        string parser,
        int rowIndex,
        IReadOnlyList<string> headers,
        IReadOnlyList<string> values,
        IReadOnlyList<string> expectedHeaders,
        string?[][]? expectedTextRows,
        object?[][]? expectedObjectRows)
    {
        if (rowIndex == 0 && !headers.SequenceEqual(expectedHeaders, StringComparer.Ordinal))
        {
            throw new InvalidOperationException($"{method} produced a different CSV header ({parser}).");
        }

        if (values.Count != expectedHeaders.Count)
        {
            throw new InvalidOperationException(
                $"{method} row {rowIndex + 1} contains {values.Count} fields; expected {expectedHeaders.Count} ({parser}).");
        }

        if (expectedTextRows != null)
        {
            if (rowIndex >= expectedTextRows.Length)
            {
                throw new InvalidOperationException($"{method} produced more than {expectedTextRows.Length} data rows ({parser}).");
            }

            var expected = expectedTextRows[rowIndex];
            for (var fieldIndex = 0; fieldIndex < expected.Length; fieldIndex++)
            {
                var expectedValue = expected[fieldIndex] ?? string.Empty;
                if (!string.Equals(values[fieldIndex], expectedValue, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"{method} row {rowIndex + 1}, field {fieldIndex + 1} did not round-trip the prepared text value ({parser}).");
                }
            }
        }
        else if (expectedObjectRows != null)
        {
            if (rowIndex >= expectedObjectRows.Length)
            {
                throw new InvalidOperationException($"{method} produced more than {expectedObjectRows.Length} data rows ({parser}).");
            }

            var expected = expectedObjectRows[rowIndex];
            for (var fieldIndex = 0; fieldIndex < expected.Length; fieldIndex++)
            {
                if (!SemanticallyEquals(values[fieldIndex], expected[fieldIndex]))
                {
                    throw new InvalidOperationException(
                        $"{method} row {rowIndex + 1}, field {fieldIndex + 1} did not round-trip the typed value '{expected[fieldIndex]}' ({parser}).");
                }
            }
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
