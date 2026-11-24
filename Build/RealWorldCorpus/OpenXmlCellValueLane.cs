using System.Text.RegularExpressions;
using System.Data.Common;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.RealWorldCorpus;

/// <summary>
/// Compares the cell values OfficeIMO's forward-only workbook reader returns with the cell values the Open XML SDK
/// reads, worksheet by worksheet. Both sides are normalized to a type prefix and an invariant text: shared and inline
/// strings to their text (without phonetic runs), booleans to TRUE/FALSE, numbers to their round-trip double, error
/// cells to their error text. Date formats are not applied, so dates compare as serial numbers, and ISO 8601 date
/// cells (<c>t="d"</c>) compare as the serial of their date. The SDK leaves <c>_xHHHH_</c> character escapes in string
/// text, so the reference decodes them, as spreadsheet applications do.
/// </summary>
internal static class OpenXmlCellValueLane {
    internal const string Name = "excel-cell-values";
    internal const long MaxCells = 1_000_000L;
    internal const long MaxWorkbookCells = 4_000_000L;

    private sealed class ComparisonBudgetException : Exception {
    }

    internal static CorpusOpenXmlLane Compare(byte[] snapshot, bool includeText, Func<Exception, string> typeName, Func<Exception, string?> message) {
        var lane = new CorpusOpenXmlLane { Name = Name };
        Workbook? expected = null;
        try {
            expected = ReadWithSdk(snapshot);
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            lane.SdkException = typeName(exception);
            lane.SdkExceptionMessage = message(exception);
        }
        if (expected != null && (expected.Extents.Any(extent => extent.Area > MaxCells) || expected.Extents.Sum(extent => extent.Area) > MaxWorkbookCells)) {
            lane.Outcome = CorpusOpenXmlOutcomes.Unavailable;
            lane.Note = "the used extents are over the comparison budget of " + MaxCells.ToString("N0", CultureInfo.InvariantCulture)
                + " cells per worksheet and " + MaxWorkbookCells.ToString("N0", CultureInfo.InvariantCulture) + " per workbook";
            return lane;
        }
        Workbook? actual = null;
        try {
            actual = ReadWithOfficeIMO(snapshot, expected);
        } catch (ComparisonBudgetException) {
            // The comparison's own budget, reached only when the reference could not bound the read in advance.
            lane.Outcome = CorpusOpenXmlOutcomes.Unavailable;
            lane.Note = "the workbook reader returned more than " + MaxWorkbookCells.ToString("N0", CultureInfo.InvariantCulture) + " cells";
            return lane;
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            lane.OurException = typeName(exception);
            lane.OurExceptionMessage = message(exception);
        }
        if (expected == null || actual == null) {
            lane.Outcome = expected != null ? CorpusOpenXmlOutcomes.OurError : actual != null ? CorpusOpenXmlOutcomes.SdkError : CorpusOpenXmlOutcomes.BothReject;
            return lane;
        }
        lane.ExpectedCount = expected.Cells.Count;
        lane.ActualCount = actual.Cells.Count;
        List<CorpusOpenXmlDifference> differences = Diff(expected, actual, includeText);
        lane.Outcome = differences.Count == 0 ? CorpusOpenXmlOutcomes.Match : CorpusOpenXmlOutcomes.Mismatch;
        lane.DifferenceCount = differences.Count;
        lane.Classes = differences.Select(difference => difference.Class).Distinct(StringComparer.Ordinal).OrderBy(name => name, StringComparer.Ordinal).ToArray();
        lane.Differences = differences.Take(5).ToArray();
        return lane;
    }

    private sealed class Workbook {
        internal List<string> Sheets { get; } = new();
        // Sheet position, row and column (1-based) to the normalized value; empty cells are absent.
        internal SortedDictionary<(int Sheet, int Row, int Column), string> Cells { get; } = new();
        // The rectangle every cell element of each worksheet falls in, whether or not it holds a value.
        internal List<Extent> Extents { get; } = new();
    }

    private sealed class Extent {
        internal int MinRow { get; set; } = int.MaxValue;
        internal int MinColumn { get; set; } = int.MaxValue;
        internal int MaxRow { get; set; }
        internal int MaxColumn { get; set; }
        internal bool IsEmpty => MaxRow == 0;
        internal long Area => IsEmpty ? 0 : (long)(MaxRow - MinRow + 1) * (MaxColumn - MinColumn + 1);

        internal void Add(int row, int column) {
            MinRow = Math.Min(MinRow, row);
            MinColumn = Math.Min(MinColumn, column);
            MaxRow = Math.Max(MaxRow, row);
            MaxColumn = Math.Max(MaxColumn, column);
        }
    }

    private static Workbook ReadWithSdk(byte[] snapshot) {
        using var stream = new MemoryStream(snapshot, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
        WorkbookPart workbookPart = document.WorkbookPart ?? throw new InvalidDataException("The package has no workbook part.");
        S.SharedStringItem[] sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<S.SharedStringItem>().ToArray()
            ?? Array.Empty<S.SharedStringItem>();
        var workbook = new Workbook();
        foreach (S.Sheet sheet in workbookPart.Workbook?.Sheets?.Elements<S.Sheet>() ?? Enumerable.Empty<S.Sheet>()) {
            if (sheet.Id?.Value is not string id || workbookPart.GetPartById(id) is not WorksheetPart worksheetPart) continue;
            int sheetIndex = workbook.Sheets.Count;
            workbook.Sheets.Add(sheet.Name?.Value ?? string.Empty);
            var extent = new Extent();
            workbook.Extents.Add(extent);
            int rowNumber = 0;
            foreach (S.Row row in worksheetPart.Worksheet?.GetFirstChild<S.SheetData>()?.Elements<S.Row>() ?? Enumerable.Empty<S.Row>()) {
                // A row without r takes its number from its first cell reference, as Excel and OfficeIMO's reader do,
                // and only then from its position.
                rowNumber = row.RowIndex?.Value is uint index ? (int)index
                    : row.Elements<S.Cell>().Select(cell => cell.CellReference?.Value).FirstOrDefault(reference => reference != null) is string first
                      && TryParseRow(first, out int referencedRow) ? referencedRow
                    : rowNumber + 1;
                int columnNumber = 0;
                foreach (S.Cell cell in row.Elements<S.Cell>()) {
                    columnNumber = cell.CellReference?.Value is string reference && TryParseColumn(reference, out int column) ? column : columnNumber + 1;
                    extent.Add(rowNumber, columnNumber);
                    string? value = SdkValue(cell, sharedStrings);
                    if (value == null) continue;
                    workbook.Cells[(sheetIndex, rowNumber, columnNumber)] = value;
                }
            }
        }
        return workbook;
    }

    private static string? SdkValue(S.Cell cell, S.SharedStringItem[] sharedStrings) {
        S.CellValues type = cell.DataType?.Value ?? S.CellValues.Number;
        string? raw = cell.CellValue?.Text;
        if (type == S.CellValues.InlineString) {
            S.InlineString? inline = cell.InlineString;
            return inline == null ? raw == null ? null : "s:" + Unescape(raw) : "s:" + RichText(inline);
        }
        if (raw == null) return null;
        if (type == S.CellValues.SharedString) {
            if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) || index < 0 || index >= sharedStrings.Length) {
                throw new InvalidDataException("A cell refers to a shared string that does not exist.");
            }
            return "s:" + RichText(sharedStrings[index]);
        }
        if (type == S.CellValues.Boolean) return "b:" + (raw.Trim() is "1" or "true" ? "TRUE" : "FALSE");
        if (type == S.CellValues.String || type == S.CellValues.Error) return "s:" + Unescape(raw);
        if (type == S.CellValues.Date) {
            // A time without a date is a fraction of a day, not a time today. It has no date to apply an offset to,
            // so the offset is ignored, and a fraction is kept to the seven digits a tick holds.
            Match timeOnly = TimeOnlyPattern.Match(raw.Trim());
            if (timeOnly.Success) {
                string fraction = (timeOnly.Groups[4].Value + "0000000").Substring(0, 7);
                TimeSpan timeOfDay = new TimeSpan(
                    int.Parse(timeOnly.Groups[1].Value, CultureInfo.InvariantCulture),
                    int.Parse(timeOnly.Groups[2].Value, CultureInfo.InvariantCulture),
                    timeOnly.Groups[3].Success ? int.Parse(timeOnly.Groups[3].Value, CultureInfo.InvariantCulture) : 0)
                    + TimeSpan.FromTicks(long.Parse(fraction, CultureInfo.InvariantCulture));
                return "n:" + timeOfDay.TotalDays.ToString("R", CultureInfo.InvariantCulture);
            }
            // Text without an offset is a wall-clock time. Text with an offset is converted to the runner's local time,
            // as the Excel reader does, so the comparison gives the same result in every timezone.
            if (!DateTimeOffset.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out DateTimeOffset parsed)) return "d:" + raw;
            DateTime date = HasExplicitOffset(raw) ? parsed.LocalDateTime : parsed.DateTime;
            return "n:" + date.ToOADate().ToString("R", CultureInfo.InvariantCulture);
        }
        return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
            ? "n:" + number.ToString("R", CultureInfo.InvariantCulture)
            : "s:" + raw;
    }

    private static readonly Regex TimeOnlyPattern = new(
        @"^[Tt]?([01][0-9]|2[0-3]):([0-5][0-9])(?::([0-5][0-9])(?:[.,]([0-9]+))?)?(?:[Zz]|[+-][0-9]{2}(?::?[0-9]{2})?)?$",
        RegexOptions.CultureInvariant);

    // An ISO 8601 value carries an offset when it ends with Z or with +hh:mm or -hh:mm after the time.
    private static bool HasExplicitOffset(string raw) {
        string text = raw.Trim();
        int time = text.IndexOf('T');
        if (time < 0) return false;
        if (text.EndsWith("Z", StringComparison.OrdinalIgnoreCase)) return true;
        return text.IndexOfAny(new[] { '+', '-' }, time) > 0;
    }

    // The text of a string item: its own t element, or its runs' t elements. Phonetic runs are not cell text.
    private static string RichText(OpenXmlElement item) {
        if (item.GetFirstChild<S.Text>() is S.Text text) return Unescape(text.Text);
        return Unescape(string.Concat(item.Elements<S.Run>().Select(run => run.Text?.Text ?? string.Empty)));
    }

    // ST_Xstring escapes a character as _xHHHH_ (ECMA-376 Part 1, 22.9.2.19).
    private static string Unescape(string value) {
        if (value.IndexOf("_x", StringComparison.Ordinal) < 0) return value;
        var text = new System.Text.StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '_' && index + 6 < value.Length && value[index + 1] == 'x' && value[index + 6] == '_'
                && int.TryParse(value.AsSpan(index + 2, 4), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out int code)) {
                text.Append((char)code);
                index += 6;
            } else {
                text.Append(value[index]);
            }
        }
        return text.ToString();
    }

    // The reader runs with its defaults apart from the header row and date conversion, so it takes the native
    // workbook path a caller gets. That path exposes each worksheet's used range, so its rows and columns are placed
    // at the top-left cell of the worksheet's cell elements as the SDK reads them; a reader that starts its used range
    // elsewhere shows up as shifted cells.
    private static Workbook ReadWithOfficeIMO(byte[] snapshot, Workbook? reference) {
        var options = new ExcelReaderOptions {
            HasHeaderRow = false,
            TreatDatesUsingNumberFormat = false
        };
        var workbook = new Workbook();
        using var stream = new MemoryStream(snapshot, writable: false);
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(stream, options);
        foreach (string name in reader.SheetNames) workbook.Sheets.Add(name);
        long cells = 0;
        do {
            int sheetIndex = reader.CurrentSheetIndex;
            Extent? origin = reference != null && sheetIndex < reference.Extents.Count && !reference.Extents[sheetIndex].IsEmpty
                ? reference.Extents[sheetIndex]
                : null;
            int row = (origin?.MinRow ?? 1) - 1;
            int firstColumn = origin?.MinColumn ?? 1;
            while (reader.Read()) {
                row++;
                cells += reader.FieldCount;
                if (cells > MaxWorkbookCells) throw new ComparisonBudgetException();
                for (int column = 0; column < reader.FieldCount; column++) {
                    string? value = OurValue(reader, column);
                    if (value != null) workbook.Cells[(sheetIndex, row, firstColumn + column)] = value;
                }
            }
        } while (reader.NextResult());
        return workbook;
    }

    private static string? OurValue(DbDataReader reader, int ordinal) {
        if (reader.IsDBNull(ordinal)) return null;
        object value = reader.GetValue(ordinal);
        return value switch {
            string text => text.Length == 0 ? null : "s:" + text,
            bool flag => "b:" + (flag ? "TRUE" : "FALSE"),
            double number => "n:" + number.ToString("R", CultureInfo.InvariantCulture),
            float number => "n:" + ((double)number).ToString("R", CultureInfo.InvariantCulture),
            decimal number => "n:" + ((double)number).ToString("R", CultureInfo.InvariantCulture),
            int or long or short => "n:" + Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture),
            DateTime date => "n:" + date.ToOADate().ToString("R", CultureInfo.InvariantCulture),
            _ => "o:" + Convert.ToString(value, CultureInfo.InvariantCulture)
        };
    }

    private static List<CorpusOpenXmlDifference> Diff(Workbook expected, Workbook actual, bool includeText) {
        var differences = new List<CorpusOpenXmlDifference>();
        if (!expected.Sheets.SequenceEqual(actual.Sheets, StringComparer.Ordinal)) {
            differences.Add(new CorpusOpenXmlDifference {
                Class = "sheet-list",
                ExpectedLength = expected.Sheets.Count,
                ActualLength = actual.Sheets.Count,
                Expected = includeText ? string.Join(" | ", expected.Sheets) : null,
                Actual = includeText ? string.Join(" | ", actual.Sheets) : null
            });
        }
        foreach (var key in expected.Cells.Keys.Union(actual.Cells.Keys).OrderBy(key => key)) {
            expected.Cells.TryGetValue(key, out string? left);
            actual.Cells.TryGetValue(key, out string? right);
            // An empty string cell and an absent cell carry the same value.
            if (left == "s:") left = null;
            if (string.Equals(left, right, StringComparison.Ordinal)) continue;
            string cellClass = left == null ? "extra-cell"
                : right == null ? "missing-cell"
                : left[0] != right[0] ? "type-" + Kind(left) + "-as-" + Kind(right)
                : Kind(left) + "-value";
            differences.Add(new CorpusOpenXmlDifference {
                Class = cellClass,
                Subject = SheetLabel(expected, actual, key.Sheet) + "!" + ColumnName(key.Column) + key.Row.ToString(CultureInfo.InvariantCulture),
                Expected = includeText ? left : null,
                Actual = includeText ? right : null
            });
        }
        return differences;
    }

    private static string Kind(string value) => value[0] switch {
        's' => "text",
        'n' => "number",
        'b' => "boolean",
        'd' => "date",
        _ => "other"
    };

    private static string SheetLabel(Workbook expected, Workbook actual, int sheet) =>
        sheet < expected.Sheets.Count ? expected.Sheets[sheet] : sheet < actual.Sheets.Count ? actual.Sheets[sheet] : "#" + sheet;

    private static bool TryParseRow(string reference, out int row) {
        int index = 0;
        while (index < reference.Length && char.IsAsciiLetter(reference[index])) index++;
        return int.TryParse(reference.AsSpan(index), NumberStyles.None, CultureInfo.InvariantCulture, out row) && row is > 0 and <= 1_048_576;
    }

    private static bool TryParseColumn(string reference, out int column) {
        column = 0;
        int index = 0;
        while (index < reference.Length && char.IsAsciiLetter(reference[index])) {
            column = checked(column * 26 + (char.ToUpperInvariant(reference[index]) - 'A' + 1));
            index++;
        }
        return index > 0 && column <= 16_384;
    }

    private static string ColumnName(int column) {
        var name = new System.Text.StringBuilder();
        while (column > 0) {
            int remainder = (column - 1) % 26;
            name.Insert(0, (char)('A' + remainder));
            column = (column - 1) / 26;
        }
        return name.ToString();
    }
}
