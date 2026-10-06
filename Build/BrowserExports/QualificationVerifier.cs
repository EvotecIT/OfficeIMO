using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.CSV;

internal static class QualificationVerifier {
    private static readonly CultureInfo Culture = CultureInfo.InvariantCulture;
    internal static string LongText => new string('a', 32766) + "🧪" + string.Concat(Enumerable.Repeat("Łódź\r\nשלום_x0041_", 2500));
    internal static object VerifyFailure(string path, ExportQualification.Case spec) {
        if (!spec.CancelAfterRows.HasValue && !spec.HangSink && !spec.ResourceLimit && !spec.PendingPage) throw new InvalidDataException("Unexpected rejection.");
        if (spec.Format == "xlsx") {
            try { using var zip = ZipFile.OpenRead(path); throw new InvalidDataException("Failed XLSX finalized a ZIP archive."); }
            catch (InvalidDataException error) when (!error.Message.StartsWith("Failed XLSX", StringComparison.Ordinal)) { }
        }
        return new { expectedFailure = true, partialBytesOwnedByCaller = new FileInfo(path).Length };
    }
    internal static object Verify(string path, ExportQualification.Case spec) {
        if (spec.Format == "csv") {
            using var reader = CsvDocument.OpenDataReader(path, new CsvLoadOptions { MaxInputBytes = 2L * 1024 * 1024 * 1024, MaxDecompressedBytes = 2L * 1024 * 1024 * 1024 });
            if (reader.FieldCount != spec.Columns) throw new InvalidDataException("CSV field count differs.");
            for (int c = 0; c < spec.Columns; c++) Require(reader.GetName(c) == "Column " + c, "CSV header differs.");
            int count = 0;
            while (reader.Read()) {
                if (count >= spec.Rows) throw new InvalidDataException("Extra CSV data row.");
                for (int c = 0; c < spec.Columns; c++) Require(reader.GetString(c) == Expected(spec, count, c, false), $"CSV value differs at {count}/{c}.");
                count++;
            }
            Require(count == spec.Rows, "CSV row count differs.");
            return new { rows = count, cells = (long)count * spec.Columns, allValues = true, reader = "OfficeIMO.CSV.OpenDataReader" };
        }
        using var archive = ZipFile.OpenRead(path);
        if (spec.Fallback) Require(archive.Entries.All(e => e.Length == e.CompressedLength), "Compression fallback did not store entries.");
        var sheet = archive.GetEntry("xl/worksheets/sheet1.xml") ?? throw new InvalidDataException("Missing worksheet.");
        int headerRows = spec.Styled ? 2 : 1, row = 0, column = 0;
        string[] letters = Enumerable.Range(1, spec.Columns).Select(ColumnName).ToArray();
        using (Stream stream = sheet.Open()) using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit })) {
            while (reader.Read()) {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.LocalName == "row") { Require(column == 0 || column == spec.Columns, "Worksheet row width differs."); row++; column = 0; Require(reader.GetAttribute("r") == row.ToString(Culture), "Worksheet row address differs."); }
                else if (reader.LocalName == "c") {
                    Require(column < spec.Columns && reader.GetAttribute("r") == letters[column] + row, "Cell address differs.");
                    string? type = reader.GetAttribute("t"); string value = ReadCell(reader);
                    if (row > headerRows && row <= spec.Rows + headerRows) {
                        int index = row - headerRows - 1;
                        if (spec.LongText && index == 100 && column == 1) Require(value.StartsWith(new string('a', 100), StringComparison.Ordinal) && value.Contains("[Full text:", StringComparison.Ordinal), "Long text preview missing.");
                        else Require(value == Expected(spec, index, column, true), $"XLSX value differs at {index}/{column}.");
                        Require(type == (column % 4 == 1 ? "inlineStr" : column % 4 == 2 ? "b" : null), "XLSX typed cell differs.");
                    } else if (row == spec.Rows + headerRows + 1 && spec.Styled && column % 4 == 0) {
                        double sum = (double)spec.Columns * spec.Rows * (spec.Rows - 1) / 2 + (double)column * spec.Rows;
                        Require(double.Parse(value, Culture) == sum, "Cached total differs.");
                    }
                    column++;
                }
            }
        }
        Require(row == spec.Rows + headerRows + (spec.Styled ? 1 : 0) && column == spec.Columns, "XLSX row count differs.");
        if (spec.LongText) {
            using Stream overflow = archive.GetEntry("xl/worksheets/sheet2.xml")!.Open();
            using XmlReader reader = XmlReader.Create(overflow); var text = new StringBuilder();
            while (reader.Read()) if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c" && reader.GetAttribute("r")!.StartsWith("D", StringComparison.Ordinal) && reader.GetAttribute("r") != "D1") {
                string chunk = ReadCell(reader); Require(chunk.Length <= 32767 && !char.IsHighSurrogate(chunk[^1]), "Invalid preservation boundary."); text.Append(chunk);
            }
            Require(text.ToString() == LongText, "Full text reconstruction differs.");
        }
        if (spec.Styled) {
            using var tableStream = archive.GetEntry("xl/tables/table1.xml")!.Open(); XDocument table = XDocument.Load(tableStream);
            Require(table.Root!.Attribute("ref")!.Value == "A2:" + letters[^1] + (spec.Rows + 3), "Grouped table range differs.");
            using var styleStream = archive.GetEntry("xl/styles.xml")!.Open(); XDocument styles = XDocument.Load(styleStream);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            Require(styles.Root!.Element(ns + "cellXfs")!.Elements().Count() < 32, "Repeated presentation inflated style registry.");
        }
        // Full schema + both reader proof is bounded to smaller artifacts; large files above are checked cell by cell without a DOM.
        if (spec.Rows <= 10000) WorkbookVerifier.Verify(path, 64L * 1024 * 1024);
        return new { rows = spec.Rows, cells = (long)spec.Rows * spec.Columns, allValues = true, preservation = spec.LongText, schema = spec.Rows <= 10000, validator = "streamed independent ZIP/XML and expected values" };
    }
    private static string ReadCell(XmlReader outer) {
        using XmlReader reader = outer.ReadSubtree(); var result = new StringBuilder();
        while (reader.Read()) if (reader.NodeType == XmlNodeType.Element && reader.LocalName is "t" or "v") result.Append(reader.ReadElementContentAsString());
        return result.ToString();
    }
    private static string Expected(ExportQualification.Case spec, int row, int column, bool excel) {
        if (spec.LongText && row == 100 && column == 1) return LongText;
        return (column % 4) switch {
            0 => ((long)row * spec.Columns + column).ToString(Culture),
            1 => spec.Unique ? "Unique " + row + ": Łódź🧪" : "Site " + row % 8,
            2 => excel ? row % 2 == 0 ? "1" : "0" : row % 2 == 0 ? "True" : "False",
            _ => excel ? new DateTime(2026, 1, 1 + row % 28).ToOADate().ToString(Culture) : new DateTime(2026, 1, 1 + row % 28).ToString("yyyy-MM-dd'T'00:00:00.000'Z'", Culture)
        };
    }
    private static string ColumnName(int value) { string text = ""; while (value > 0) { value--; text = (char)('A' + value % 26) + text; value /= 26; } return text; }
    private static void Require(bool value, string message) { if (!value) throw new InvalidDataException(message); }
}
