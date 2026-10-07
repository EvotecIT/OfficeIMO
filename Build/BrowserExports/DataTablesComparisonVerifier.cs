using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using OfficeIMO.CSV;

/// <summary>Every-value streaming validation of the common plain export contract.</summary>
internal static class DataTablesComparisonVerifier {
    internal static object Verify(string path, string format, int rows, int columns, bool unique, bool fullWidthScan = false, string textProfile = "unicode") {
        long cells = 0;
        if (format == "csv") {
            using var reader = CsvDocument.OpenDataReader(path, new CsvLoadOptions { MaxInputBytes = 2L * 1024 * 1024 * 1024, MaxDecompressedBytes = 2L * 1024 * 1024 * 1024 });
            Require(reader.FieldCount == columns, "CSV header width differs.");
            for (int column = 0; column < columns; column++) Require(reader.GetName(column) == "Column " + (column + 1), "CSV header differs.");
            cells += columns;
            int row = 0;
            while (reader.Read()) {
                Require(reader.FieldCount == columns, "CSV column count differs.");
                for (int column = 0; column < columns; column++) {
                    Require(reader.GetString(column) == Expected(row, column, unique, textProfile), $"CSV value differs at {row}/{column}."); cells++;
                }
                row++;
            }
            Require(row == rows, "CSV row count differs.");
        } else {
            using var archive = ZipFile.OpenRead(path);
            using var stream = (archive.GetEntry("xl/worksheets/sheet1.xml") ?? throw new InvalidDataException("Missing worksheet.")).Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
            int row = 0, column = 0, widthColumns = 0;
            double[] widths = new double[columns]; int[] longest = new int[columns];
            while (reader.Read()) {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.LocalName == "col") {
                    int first = int.Parse(reader.GetAttribute("min")!, CultureInfo.InvariantCulture), last = int.Parse(reader.GetAttribute("max")!, CultureInfo.InvariantCulture);
                    Require(first == widthColumns + 1 && last >= first && last <= columns, "XLSX column width ranges differ.");
                    double width = double.Parse(reader.GetAttribute("width")!, CultureInfo.InvariantCulture);
                    Require(fullWidthScan ? width >= 6 && width <= 54 : width == 20, "XLSX column widths violate the requested sizing contract.");
                    for (int i = first - 1; i < last; i++) widths[i] = width;
                    widthColumns = last;
                } else if (reader.LocalName == "row") {
                    Require(row == 0 || column == columns, "XLSX row width differs."); row++; column = 0;
                    Require(reader.GetAttribute("r") == row.ToString(CultureInfo.InvariantCulture), "XLSX row coordinate differs.");
                } else if (reader.LocalName == "c") {
                    using XmlReader subtree = reader.ReadSubtree(); XElement cell = XElement.Load(subtree);
                    Require(column < columns && row > 0, "Unexpected XLSX cell.");
                    Require(cell.Attribute("r")?.Value == ColumnName(column + 1) + row, "XLSX cell coordinate differs.");
                    string actual = string.Concat(cell.Descendants().Where(e => e.Name.LocalName is "t" or "v").Select(e => e.Value));
                    string expected = row == 1 ? "Column " + (column + 1) : Expected(row - 2, column, unique, textProfile);
                    longest[column] = Math.Max(longest[column], expected.EnumerateRunes().Count());
                    bool numeric = row > 1 && column % 3 != 1;
                    Require(numeric ? double.TryParse(actual, CultureInfo.InvariantCulture, out double number) && number == double.Parse(expected, CultureInfo.InvariantCulture) : actual == expected,
                        $"XLSX value differs at {row}/{column}.");
                    Require(numeric ? cell.Attribute("t") is null or { Value: "n" } : cell.Attribute("t")?.Value == "inlineStr", "XLSX cell type differs.");
                    column++; cells++;
                }
            }
            Require(row == rows + 1 && column == columns && widthColumns == columns, "XLSX final dimensions differ.");
            if (fullWidthScan) for (int i = 0; i < columns; i++)
                Require(widths[i] >= Math.Min(54, Math.Max(6, longest[i])), "XLSX width misses a value in the complete table.");
        }
        return new { rows, columns, cells, allValues = true, contract = "ordered typed numbers and literal Unicode strings; headers; " + (fullWidthScan ? "whole-table approximate width sizing" : "fixed width 20") };
    }

    private static string Expected(int row, int column, bool unique, string textProfile) => column % 3 == 0 ? ((long)row * 10 + column).ToString(CultureInfo.InvariantCulture)
        : column % 3 == 1 ? $"Łódź{(textProfile == "bmp" ? "" : " 🧪")} {(unique ? "row-" + row : "group-" + row % 100)}-column-{column}"
        : (row % 10000 + column / 100d).ToString("G", CultureInfo.InvariantCulture);
    private static string ColumnName(int column) { string name = ""; while (column > 0) { column--; name = (char)('A' + column % 26) + name; column /= 26; } return name; }
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidDataException(message); }
}
