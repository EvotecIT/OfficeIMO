using OfficeIMO.Excel;
using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;

namespace OfficeIMO.LegacyImport.Tests;

public sealed partial class EarlyBiffImportTests {
    [Fact]
    public void NativeBiff4ArrayGroupsRetainRangesCachesAndUnsupportedGroupFindingsAfterEditing() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "EarlyBiff", "formula_import_biff4.xlw");
        byte[] original = File.ReadAllBytes(path);
        using ExcelDocument document = ExcelDocument.Load(path, null, out OfficeConversionReport report);
        Assert.Equal(new[] { "Cell", "Param", "Functions", "Array", "Table" }, document.Sheets.Select(sheet => sheet.Name));
        Assert.Equal(new[] {
            "'Array'!B13", "'Array'!C13", "'Array'!B14", "'Array'!C14",
            "'Array'!B22", "'Array'!C22", "'Array'!B23", "'Array'!C23",
            "'Array'!B27", "'Array'!C27", "'Array'!B28", "'Array'!C28"
        }, report.Diagnostics.Where(finding => finding.Code == "early-biff-formula-unsupported"
            && finding.Location?.StartsWith("'Array'!", StringComparison.Ordinal) == true).Select(finding => finding.Location));

        document.Sheets[3].CellValue(40, 1, "edited");
        using var output = new MemoryStream();
        document.Save(output, new ExcelSaveOptions { AllowLoss = true });
        XDocument worksheet = ReadArrayWorksheet(output, 4);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var cells = worksheet.Descendants(ns + "c").ToDictionary(cell => (string)cell.Attribute("r")!);
        var groups = new[] {
            (FirstRow: 7, LastRow: 8, Formula: "1", Values: new[] { 1d, 1d, 1d, 1d }),
            (FirstRow: 9, LastRow: 10, Formula: "H9:I10", Values: new[] { 1d, 2d, 3d, 4d }),
            (FirstRow: 11, LastRow: 12, Formula: "$H$9:$I$10", Values: new[] { 1d, 2d, 3d, 4d }),
            (FirstRow: 18, LastRow: 19, Formula: "MINVERSE(H18:I19)+H20:I21", Values: new[] { 1d, 2d, 3d, 4d }),
            (FirstRow: 20, LastRow: 21, Formula: "MINVERSE($H$18:$I$19)+$H$20:$I$21", Values: new[] { 1d, 2d, 3d, 4d })
        };
        Assert.Equal(5, worksheet.Descendants(ns + "f").Count(formula => (string?)formula.Attribute("t") == "array"));
        var expectedReaderValues = new Dictionary<(int Row, int Column), double>();
        foreach (var group in groups) {
            XElement owner = cells["B" + group.FirstRow].Element(ns + "f")!;
            Assert.Equal("array", (string?)owner.Attribute("t"));
            Assert.Equal("B" + group.FirstRow + ":C" + group.LastRow, (string?)owner.Attribute("ref"));
            Assert.Equal(group.Formula, owner.Value);
            for (int index = 0; index < 4; index++) {
                int row = group.FirstRow + index / 2;
                int column = 2 + index % 2;
                XElement cell = cells[(column == 2 ? "B" : "C") + row];
                Assert.Equal("n", (string?)cell.Attribute("t") ?? "n");
                Assert.Equal(BitConverter.DoubleToInt64Bits(group.Values[index]),
                    BitConverter.DoubleToInt64Bits(double.Parse(cell.Element(ns + "v")!.Value, CultureInfo.InvariantCulture)));
                if (index != 0) Assert.Null(cell.Element(ns + "f"));
                expectedReaderValues.Add((row, column), group.Values[index]);
            }
        }
        foreach (int row in new[] { 13, 22, 27 }) {
            for (int index = 0; index < 4; index++) {
                XElement cell = cells[(index % 2 == 0 ? "B" : "C") + (row + index / 2)];
                Assert.Null(cell.Element(ns + "f"));
                Assert.Equal(index + 1d, double.Parse(cell.Element(ns + "v")!.Value, CultureInfo.InvariantCulture));
            }
        }
        using (var reopenedSource = new MemoryStream(output.ToArray()))
        using (ExcelDocument reopened = ExcelDocument.Load(reopenedSource)) {
            Assert.Equal("edited", reopened.Sheets[3].Cell(40, 1).GetValue<string>());
        }
        using (ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path,
            new ExcelReaderOptions { SheetName = "Array", HasHeaderRow = false })) {
            int row = 0;
            int matched = 0;
            while (reader.Read()) {
                row++;
                for (int column = 2; column <= 3; column++) {
                    if (!expectedReaderValues.TryGetValue((row, column), out double value)) continue;
                    Assert.Equal(BitConverter.DoubleToInt64Bits(value), BitConverter.DoubleToInt64Bits(reader.GetDouble(column - 1)));
                    matched++;
                }
            }
            Assert.Equal(20, matched);
        }
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    [Fact]
    public void Biff4ArrayGroupRetainsOriginalCachesWithoutEvaluationAndUsesTheProjectedFormulaReader() {
        byte[] source = BuildArrayWorksheet(complete: true);
        using var input = new MemoryStream(source);
        using ExcelDocument document = ExcelDocument.Load(input, null, out OfficeConversionReport report);
        Assert.Empty(report.Diagnostics);
        using var output = new MemoryStream();
        document.Save(output);
        XDocument worksheet = ReadArrayWorksheet(output, 1);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var cells = worksheet.Descendants(ns + "c").ToArray();
        Assert.Equal(new[] { "77", "78", "79", "80" }, cells.Select(cell => cell.Element(ns + "v")!.Value));
        XElement formula = Assert.Single(worksheet.Descendants(ns + "f"));
        Assert.Equal("1", formula.Value);
        Assert.Equal("array", (string?)formula.Attribute("t"));
        Assert.Equal("A1:B2", (string?)formula.Attribute("ref"));
        using var readerSource = new MemoryStream(source);
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(readerSource,
            new ExcelReaderOptions { HasHeaderRow = false, UseCachedFormulaResult = false });
        Assert.True(reader.Read());
        Assert.Equal("1", reader.GetString(0));
        Assert.Equal(78d, reader.GetDouble(1));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IncompleteBiff4ArrayGroupsRetainCachesAndDoNotEraseAnUnrelatedFormula(bool outsideMember) {
        using var input = new MemoryStream(BuildArrayWorksheet(complete: false, outsideMember));
        using ExcelDocument document = ExcelDocument.Load(input, null, out OfficeConversionReport report);
        Assert.Equal(outsideMember ? new[] { "'Sheet1'!A1", "'Sheet1'!B1", "'Sheet1'!A2", "'Sheet1'!A3" }
            : new[] { "'Sheet1'!A1", "'Sheet1'!B1", "'Sheet1'!A2" },
            report.Diagnostics.Where(finding => finding.Code == "early-biff-formula-unsupported").Select(finding => finding.Location));
        using var output = new MemoryStream();
        document.Save(output, new ExcelSaveOptions { AllowLoss = true });
        XDocument worksheet = ReadArrayWorksheet(output, 1);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var cells = worksheet.Descendants(ns + "c").ToDictionary(cell => (string)cell.Attribute("r")!);
        Assert.Equal("77", cells["A1"].Element(ns + "v")!.Value);
        Assert.Equal("78", cells["B1"].Element(ns + "v")!.Value);
        Assert.Equal("79", cells["A2"].Element(ns + "v")!.Value);
        Assert.Equal("99", cells["B2"].Element(ns + "v")!.Value);
        Assert.Equal("2", Assert.Single(worksheet.Descendants(ns + "f")).Value);
        if (outsideMember) Assert.Equal("81", cells["A3"].Element(ns + "v")!.Value);
    }

    [Theory]
    [InlineData(0, "Truncated early BIFF array formula.")]
    [InlineData(1, "Invalid early BIFF array formula range.")]
    [InlineData(2, "Truncated early BIFF array formula tokens.")]
    public void Biff4ArrayDeclarationsRejectMalformedHeadersBoundsAndTokenLengths(int malformed, string expectedMessage) {
        using var source = new MemoryStream();
        using (var writer = new BinaryWriter(source, Encoding.UTF8, leaveOpen: true)) {
            Record(writer, 0x0409, record => WriteBof(record, 4, 0x10));
            Record(writer, 0x0221, record => {
                if (malformed == 0) { record.Write(new byte[9]); return; }
                record.Write((ushort)(malformed == 1 ? 1 : 0));
                record.Write((ushort)0);
                record.Write((byte)0);
                record.Write((byte)0);
                record.Write((ushort)0);
                record.Write((ushort)3);
            });
            Record(writer, 0x000A, _ => { });
        }
        source.Position = 0;
        Assert.Equal(expectedMessage, Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(source)).Message);
    }

    private static XDocument ReadArrayWorksheet(MemoryStream output, int index) {
        output.Position = 0;
        using var package = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true);
        using Stream part = package.GetEntry($"xl/worksheets/sheet{index}.xml")!.Open();
        return XDocument.Load(part);
    }

    private static byte[] BuildArrayWorksheet(bool complete, bool outsideMember = false) {
        using var bytes = new MemoryStream();
        using var writer = new BinaryWriter(bytes);
        Record(writer, 0x0409, record => WriteBof(record, 4, 0x10));
        for (int index = 0; index < (complete ? 4 : 3); index++) {
            WriteFormula(index / 2, index % 2, 77d + index, new byte[] { 1, 0, 0, 0, 0 });
            if (index != 0) continue;
            Record(writer, 0x0221, record => {
                record.Write((ushort)0);
                record.Write((ushort)1);
                record.Write((byte)0);
                record.Write((byte)1);
                record.Write((ushort)0);
                record.Write((ushort)3);
                record.Write(new byte[] { 0x1E, 1, 0 });
            });
        }
        if (!complete) WriteFormula(1, 1, 99d, new byte[] { 0x1E, 2, 0 });
        if (outsideMember) WriteFormula(2, 0, 81d, new byte[] { 1, 0, 0, 0, 0 });
        Record(writer, 0x000A, _ => { });
        return bytes.ToArray();

        void WriteFormula(int row, int column, double cached, byte[] tokens) => Record(writer, 0x0406, record => {
            record.Write((ushort)row);
            record.Write((ushort)column);
            record.Write((ushort)0);
            record.Write(cached);
            record.Write((ushort)0);
            record.Write((ushort)tokens.Length);
            record.Write(tokens);
        });
    }
}
