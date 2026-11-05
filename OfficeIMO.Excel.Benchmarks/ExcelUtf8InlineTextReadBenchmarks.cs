using System.IO.Compression;
using System.Text;
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Reads unique multi-line inline strings from a UTF-8 worksheet through the public data
/// reader. CRLF text exercises literal line-ending normalization; LF is the control and
/// delivers identical strings. ImplicitRows omits row references while each cell keeps its
/// reference, so every row index is inferred from the row's first cell. Data starts in
/// column B. Setup checks every delivered value. This is an OfficeIMO before/after lane.
/// </summary>
[MemoryDiagnoser]
public class ExcelUtf8InlineTextReadBenchmarks {
    private string _path = string.Empty;
    private long _expected;

    [Params(20000)]
    public int RowCount { get; set; }

    [Params("LF", "CRLF")]
    public string LineEndings { get; set; } = "CRLF";

    [Params("Explicit", "ImplicitRows")]
    public string Coordinates { get; set; } = "Explicit";

    [GlobalSetup]
    public void Setup() {
        if (LineEndings is not ("LF" or "CRLF")) throw new ArgumentOutOfRangeException(nameof(LineEndings));
        if (Coordinates is not ("Explicit" or "ImplicitRows")) throw new ArgumentOutOfRangeException(nameof(Coordinates));
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_DATA") ?? Path.GetTempPath();
        Directory.CreateDirectory(root);
        _path = Path.Combine(root, $"utf8-inline-text-{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(_path)) {
                document.AddSheet("Data").Cell(1, 2).SetValue("Text");
                document.Save();
            }
            using (var package = ZipFile.Open(_path, ZipArchiveMode.Update)) {
                const string name = "xl/worksheets/sheet1.xml";
                ZipArchiveEntry entry = package.GetEntry(name) ?? throw new InvalidDataException("Worksheet is missing.");
                entry.Delete();
                using Stream output = package.CreateEntry(name, CompressionLevel.Fastest).Open();
                byte[] xml = new UTF8Encoding(false).GetBytes(CreateWorksheetXml());
                output.Write(xml, 0, xml.Length);
            }
            _expected = ReadCore(validate: true);
        } catch {
            Cleanup();
            throw;
        }
    }

    [Benchmark]
    public long Read() {
        long result = ReadCore(validate: false);
        if (result != _expected) throw new InvalidDataException("The inline text checksum changed.");
        return result;
    }

    [GlobalCleanup]
    public void Cleanup() {
        if (_path.Length != 0) File.Delete(_path);
    }

    private long ReadCore(bool validate) {
        using var reader = ExcelDocument.OpenReader(_path);
        if (reader.FieldCount != 2 || reader.GetName(0) != "Text" || reader.GetName(1) != "Index") {
            throw new InvalidDataException("Inline text headers differ.");
        }
        long checksum = 0;
        int row = 0;
        while (reader.Read()) {
            string text = reader.GetString(0);
            int index = reader.GetInt32(1);
            if (validate && (index != row || !string.Equals(text, Text(row, "\n"), StringComparison.Ordinal))) {
                throw new InvalidDataException($"Inline text row {row + 2} differs.");
            }
            checksum += text.Length * 31L + text[text.Length - 1] + index;
            row++;
        }
        if (row != RowCount) throw new InvalidDataException("Inline text row count differs.");
        return checksum;
    }

    private string CreateWorksheetXml() {
        string lineEnding = LineEndings == "CRLF" ? "\r\n" : "\n";
        bool explicitRows = Coordinates == "Explicit";
        var xml = new StringBuilder(RowCount * 120);
        xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
            .Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">")
            .Append("<dimension ref=\"B1:C").Append(RowCount + 1).Append("\"/><sheetData>");
        AppendRow(xml, 1, explicitRows, "<c r=\"B1\" t=\"inlineStr\"><is><t>Text</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>Index</t></is></c>");
        for (int row = 0; row < RowCount; row++) {
            int sheetRow = row + 2;
            AppendRow(xml, sheetRow, explicitRows,
                $"<c r=\"B{sheetRow}\" t=\"inlineStr\"><is><t>{Text(row, lineEnding)}</t></is></c><c r=\"C{sheetRow}\"><v>{row}</v></c>");
        }
        return xml.Append("</sheetData></worksheet>").ToString();
    }

    private static void AppendRow(StringBuilder xml, int sheetRow, bool explicitRows, string cells) {
        xml.Append(explicitRows ? $"<row r=\"{sheetRow}\">" : "<row>").Append(cells).Append("</row>");
    }

    private static string Text(int row, string lineEnding) =>
        $"Row {row} first line{lineEnding}second line {row}{lineEnding}third";
}
