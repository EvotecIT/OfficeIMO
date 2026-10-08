using System.Globalization;
using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;

internal static class WorkbookVerifier {
    internal static void Require(bool value, string message) { if (!value) throw new InvalidDataException(message); }

    internal static void Verify(string path, long? maxCharactersInPart = null) {
        OfficeIMO.TestAssets.JavaScriptWorkbookContract.Verify(path, maxCharactersInPart: maxCharactersInPart);
        using SpreadsheetDocument sdk = SpreadsheetDocument.Open(path, false);
        var errors = new OpenXmlValidator().Validate(sdk).Take(8).ToArray();
        Require(errors.Length == 0, Path.GetFileName(path) + ": " + string.Join("; ", errors.Select(e => e.Description)));
        // Exercise both the editable model and the public streaming reader.
        using var document = ExcelDocument.Load(path, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        string file = Path.GetFileNameWithoutExtension(path);
        if (file.StartsWith("rich-", StringComparison.Ordinal)) {
            var names = sdk.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Select(s => s.Name!.Value).ToArray();
            Require(names.SequenceEqual(new[] { "Domain controllers", "domain controllers (2)", "Bad_______", "Sheet", "History_", new string('a', 30) }), "Sheet names differ.");
            object?[,] values = Read(path, names[0]!, "A1:E5");
            Require(Equals(values[1, 0], "DC<&\"'\r\n\t🧪שלום"), "Unicode/XML/newline round trip failed: " + values[1, 0]);
            Require(Equals(values[1, 1], "Łódź") && Equals(values[1, 4], true) && Equals(values[2, 4], false), "String/boolean types differ.");
            Require(Convert.ToDouble(values[1, 3], CultureInfo.InvariantCulture) == 12.5 && Convert.ToDouble(values[2, 3], CultureInfo.InvariantCulture) == -1.25, "Numbers differ.");
            Require(Equals(values[2, 0], "=literal"), "Formula-like XLSX value changed.");
            Require(Equals(values[3, 0], "_x0041_"), "Literal OOXML escape text changed: " + values[3, 0]);
            Require(values[3, 3] is null && values[4, 3] is null, "Non-finite numbers must be empty.");
            Require(values[1, 2] is DateTime d && d == new DateTime(2026, 10, 5, 12, 34, 56), "Date type/value differs.");
            Require(values[2, 2] is DateTime early && early == new DateTime(1900, 2, 28, 12, 0, 0), "Early 1900 date differs.");
            Require(values[3, 2] is DateTime march && march == new DateTime(1900, 3, 1), "1900 leap-year boundary differs.");
            Require(values[4, 2] is DateTime january && january == new DateTime(1900, 1, 1), "1900 epoch differs.");
            Worksheet worksheet = sdk.WorkbookPart.WorksheetParts.First().Worksheet!;
            Require(worksheet.GetFirstChild<Columns>()!.Elements<Column>().First().Width!.Value == 28, "Width differs.");
            Require(worksheet.Descendants<Pane>().Single().VerticalSplit!.Value == 1, "Frozen pane differs.");
            Require(worksheet.GetFirstChild<AutoFilter>()!.Reference!.Value == "A1:E5", "Autofilter used range differs.");
            Stylesheet styles = sdk.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
            Require(styles.NumberingFormats!.Elements<NumberingFormat>().Any(f => f.FormatCode!.Value == "yyyy-mm-dd hh:mm"), "Date format missing.");
            Require(styles.Fills!.Descendants<ForegroundColor>().Any(f => f.Rgb!.Value == "FFD9E1F2"), "Header fill missing.");
            Require(styles.CellFormats!.Descendants<Alignment>().Any(a => a.WrapText?.Value == true && a.Horizontal?.Value == HorizontalAlignmentValues.Left), "Column alignment/wrap missing.");
            Require(styles.Fonts!.Elements<Font>().Any(f => f.Bold is not null), "Bold header missing.");
            Require(sdk.PackageProperties.Creator == "Test<&\" 🧪" && sdk.PackageProperties.Title == "Report <>&", "Properties differ.");
            Require(sdk.PackageProperties.Created?.ToUniversalTime() == new DateTime(2026, 10, 5, 10, 0, 0, DateTimeKind.Utc), "Created property differs.");
        } else if (file == "one") Require(Equals(Read(path, "One", "A1:A1")[0, 0], "one"), "One cell differs.");
        else if (file == "worker") {
            object?[,] row = Read(path, "Worker", "A2:D2");
            Require(Equals(row[0, 0], "Łódź") && Equals(row[0, 1], new DateTime(2026, 10, 5, 12, 34, 56)) &&
                Convert.ToDouble(row[0, 2], CultureInfo.InvariantCulture) == -2 && Equals(row[0, 3], true), "Worker values/types differ.");
        }
        else if (file.StartsWith("advanced-projection-", StringComparison.Ordinal)) {
            object?[,] values = Read(path, "Styled", "A2:B2");
            Require(Convert.ToDouble(values[0, 0], CultureInfo.InvariantCulture) == 123 && Convert.ToDouble(values[0, 1], CultureInfo.InvariantCulture) == 124, "Advanced Cell projection values differ.");
            Stylesheet styles = sdk.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            foreach (Cell cell in sdk.WorkbookPart.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Where(c => c.CellReference?.Value is "A2" or "B2")) {
                CellFormat format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
                Require(styles.NumberingFormats!.Elements<NumberingFormat>().Any(f => f.NumberFormatId!.Value == format.NumberFormatId!.Value && f.FormatCode!.Value == "0.000"), "Advanced Cell number format differs.");
            }
        }
        else if (file == "packed-consumer") Require(Equals(Read(path, "Packed", "A2:A2")[0, 0], "Łódź 🧪"), "Packed npm writer output differs.");
        else if (file == "packed-worker") {
            object?[,] rows = Read(path, "Data", "A2:C3");
            Require(rows.GetLength(0) == 2 && Equals(rows[0, 0], "Łódź 🧪") && Equals(rows[1, 0], "second"), "Packed worker pages differ.");
            Require(Convert.ToDouble(rows[0, 1], CultureInfo.InvariantCulture) == 12.5 && Convert.ToDouble(rows[1, 1], CultureInfo.InvariantCulture) == 7.5, "Packed worker numbers differ.");
            Require(Equals(rows[0, 2], new DateTime(2026, 10, 7)) && Equals(rows[1, 2], new DateTime(2026, 10, 7)), "Packed worker dates differ.");
        }
        else if (file == "long") Require(Equals(Read(path, "Long", "A2:A2")[0, 0], new string('a', 32765) + "🧪"), "Long string differs.");
        else if (file == "wide") Require(Equals(Read(path, "Wide", "XFD1:XFD1")[0, 0], "last"), "Maximum column differs.");
        else if (file.StartsWith("date-", StringComparison.Ordinal)) {
            DateTime expected = new DateTime(2026, 10, 5, file == "date-local" ? 14 : 12, 34, 56, 123);
            Require(Read(path, "Date", "A1:A1")[0, 0] is DateTime actual && actual == expected, "Local/UTC date differs.");
        }
        if (file is "rich-store" or "fallback" or "fallback-raw") {
            using var stream = File.OpenRead(path);
            using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
            Require(zip.Entries.All(e => e.Length == e.CompressedLength), "Stored ZIP fallback compressed entries.");
        }
    }

    internal static void VerifyScale(string path) {
        // Full SDK validation is intentionally in this explicit evidence run.
        using SpreadsheetDocument sdk = SpreadsheetDocument.Open(path, false);
        Require(!new OpenXmlValidator().Validate(sdk).Any(), "Scale workbook failed Open XML validation.");
        object?[,] last = Read(path, "Scale", "A100001:T100001");
        Require(Convert.ToDouble(last[0, 0], CultureInfo.InvariantCulture) == 1999980, "Scale last row missing.");
        Require(Equals(last[0, 19], "Site 19"), "Scale last column missing.");
    }

    private static object?[,] Read(string path, string sheet, string range) {
        using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { SheetName = sheet, A1Range = range, HasHeaderRow = false });
        var rows = new List<object?[]>();
        while (reader.Read()) {
            var values = new object[reader.FieldCount];
            reader.GetValues(values);
            rows.Add(values.Select(v => v is DBNull ? null : v).ToArray());
        }
        var result = new object?[rows.Count, reader.FieldCount];
        for (int r = 0; r < rows.Count; r++) for (int c = 0; c < rows[r].Length; c++) result[r, c] = rows[r][c];
        return result;
    }
}
