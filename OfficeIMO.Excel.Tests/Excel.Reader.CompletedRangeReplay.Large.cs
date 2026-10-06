using System.Globalization;
using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    public static IEnumerable<object[]> LargeCompletedRangeReplayCases() {
        foreach (string api in new[] { "StreamDomSequential", "StreamDomParallel", "ObjectsAutomatic", "ObjectsParallel", "DataReader", "TypedObjectsStream" })
            foreach (bool utf16 in new[] { false, true })
                foreach (bool cancellable in new[] { false, true })
                    yield return new object[] { api, utf16, cancellable };
    }

    [Theory]
    [MemberData(nameof(LargeCompletedRangeReplayCases))]
    public void Reader_LargeCompletedRangeStillReadsLaterPhysicalRows(string api, bool utf16, bool cancellable) {
        const int lastRow = 5000;
        string path = CreateCompactFastPathWorkbook();
        var xml = new StringBuilder();
        xml.Append("<?xml version=\"1.0\" encoding=\"").Append(utf16 ? "utf-16" : "utf-8").Append("\"?>");
        xml.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:A5000\"/><sheetData>");
        xml.Append("<row r=\"1\"><c r=\"A1\" t=\"str\"><v>Value</v></c></row>");
        for (int row = 2; row <= lastRow; row++)
            xml.Append("<row r=\"").Append(row).Append("\"><c r=\"A").Append(row).Append("\" t=\"str\"><v>")
                .Append(row == 2 ? "old" : row.ToString(CultureInfo.InvariantCulture)).Append("</v></c></row>");
        xml.Append("<row r=\"5001\"><c r=\"A5001\" t=\"str\"><v>outside</v></c></row>");
        xml.Append("<row r=\"2\"><c r=\"A2\" t=\"str\"><v>new</v></c></row></sheetData></worksheet>");

        try {
            ReplaceZipEntry(path, "xl/worksheets/sheet1.xml", (utf16 ? Encoding.Unicode : Encoding.UTF8).GetBytes(xml.ToString()));
            using var cancellation = new CancellationTokenSource();
            CancellationToken token = cancellable ? cancellation.Token : default;
            var options = new ExcelReadOptions();
            if (api.StartsWith("StreamDom", StringComparison.Ordinal)) options.Culture = CultureInfo.GetCultureInfo("fr-FR");
            using var owner = ExcelDocumentReader.Open(path, options);
            var sheet = owner.GetSheet("Data");
            object?[] actual;
            switch (api) {
                case "StreamDomSequential":
                case "StreamDomParallel": {
                    actual = new object?[lastRow - 1];
                    var mode = api == "StreamDomParallel" ? ExcelExecutionMode.Parallel : ExcelExecutionMode.Sequential;
                    foreach (var chunk in sheet.ReadRangeStream("A2:A5000", chunkRows: lastRow, mode: mode, ct: token))
                        for (int row = 0; row < chunk.RowCount; row++) actual[chunk.StartRow + row - 2] = chunk.Rows[row][0];
                    break;
                }
                case "ObjectsAutomatic":
                case "ObjectsParallel": {
                    var mode = api == "ObjectsParallel" ? ExcelExecutionMode.Parallel : ExcelExecutionMode.Automatic;
                    actual = sheet.ReadObjects<CompletedRangeReplayRow>("A1:A5000", mode, token).Select(row => (object?)row.Value).ToArray();
                    break;
                }
                case "TypedObjectsStream":
                    actual = sheet.ReadObjectsStream<CompletedRangeReplayRow>("A1:A5000", token).Select(row => (object?)row.Value).ToArray();
                    break;
                case "DataReader": {
                    using var reader = sheet.ReadRangeAsDataReader("A1:A5000", schemaSampleRows: 0, ct: token);
                    var values = new List<object?>();
                    while (reader.Read()) values.Add(reader.GetValue(0));
                    actual = values.ToArray();
                    break;
                }
                default: throw new InvalidOperationException(api);
            }
            Assert.Equal(lastRow - 1, actual.Length);
            Assert.Equal("new", actual[0]);
            for (int row = 3; row <= lastRow; row++) Assert.Equal(row.ToString(CultureInfo.InvariantCulture), actual[row - 2]);
        } finally {
            File.Delete(path);
        }
    }
}
