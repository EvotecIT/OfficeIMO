using System.Globalization;
using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelXStringPersistenceTests {
    private static readonly (string Value, string Xml)[] Literals = {
        ("_x0041_", "_x005F_x0041_"),
        ("_x005F_", "_x005F_x005F_"),
        ("_x0041__x0042_", "_x005F_x0041__x005F_x0042_"),
        ("Test_xfeed_xfeed_xxxx-_ffff_", "Test_x005F_xfeed_x005F_xfeed_xxxx-_ffff_"),
        ("_x00G1_ _x041_ _X0041_ 😀", "_x00G1_ _x041_ _X0041_ 😀")
    };

    [Theory]
    [InlineData("Dom")]
    [InlineData("Standard")]
    [InlineData("Rich")]
    [InlineData("Inline")]
#if NET8_0_OR_GREATER
    [InlineData("Utf8")]
#endif
    public void WritersProtectLiteralEscapeTokensWithIndependentlyExpectedXml(string route) {
        using var saved = new MemoryStream();
        if (route == "Inline" || route == "Utf8") {
            using var writer = ExcelDocument.CreateWriter(saved, new ExcelWriterOptions {
                IncludeHeaders = false, CreateTable = false, UseSharedStrings = false
            }, leaveOpen: true);
            writer.WriteTable("Text", Literals, new[] { "Value" }, (row, item) => {
#if NET8_0_OR_GREATER
                if (route == "Utf8") row.WriteUtf8(Encoding.UTF8.GetBytes(item.Value));
                else
#endif
                row.Write(item.Value);
            });
        } else {
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddSheet("Text");
            for (int index = 0; index < Literals.Length; index++) {
                if (route == "Rich") sheet.Cell(index + 1, 1).SetRichText(new ExcelRichTextRun(Literals[index].Value));
                else sheet.Cell(index + 1, 1).SetValue(Literals[index].Value);
                Assert.Equal(Literals[index].Value, sheet.Cell(index + 1, 1).GetValue<string>());
            }
            document.Save(saved, new ExcelSaveOptions { DisableFastPackageWriter = route == "Standard" });
        }
        byte[] bytes = saved.ToArray();
        using (var package = SpreadsheetDocument.Open(new MemoryStream(bytes), false)) {
            var workbook = package.WorkbookPart!;
            var strings = workbook.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().ToArray();
            Cell[] cells = workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToArray();
            Assert.Equal(Literals.Length, cells.Length);
            for (int index = 0; index < cells.Length; index++) {
                Cell cell = cells[index];
                string actual = cell.DataType?.Value == CellValues.SharedString
                    ? strings![int.Parse(cell.CellValue!.Text, CultureInfo.InvariantCulture)].InnerText
                    : cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? string.Empty;
                Assert.Equal(Literals[index].Xml, actual);
            }
        }
        using var loaded = ExcelDocument.Load(new MemoryStream(bytes));
        using var reader = ExcelDocument.OpenReader(new MemoryStream(bytes), new ExcelReaderOptions { HasHeaderRow = false });
        for (int index = 0; index < Literals.Length; index++) {
            Assert.Equal(Literals[index].Value, loaded.Sheets[0].Cell(index + 1, 1).GetValue<string>());
            Assert.True(reader.Read());
            Assert.Equal(Literals[index].Value, reader.GetString(0));
        }
        Assert.False(reader.Read());
        // Saving an imported XML payload must not encode the already escaped underscore again.
        using var resaved = ExcelDocument.Load(new MemoryStream(loaded.ToBytes()));
        for (int index = 0; index < Literals.Length; index++)
            Assert.Equal(Literals[index].Value, resaved.Sheets[0].Cell(index + 1, 1).GetValue<string>());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IndependentXStringFixtureDecodesOnceAcrossReaderTransports(bool utf16) {
        using var document = ExcelDocument.Create();
        document.AddSheet("Text").Cell(1, 1).SetValue("seed");
        using var source = new MemoryStream();
        document.Save(source);
        const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        string[] payloads = { "_x005F_x0041_", "_x0041_", "_x005F_x005F_x005F_x0041_", "_x005F_x0041_" };
        string[] expected = { "_x0041_", "A", "_x005F_x0041_", "_x0041_" };
        using (var archive = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true)) {
            var rows = new StringBuilder("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Shared</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Inline</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>Plain</t></is></c><c r=\"D1\" t=\"inlineStr\"><is><t>Rich</t></is></c></row>");
            for (int index = 0; index < payloads.Length; index++) {
                int row = index + 2;
                string rich = index == 3 ? "<r><t>_x00</t></r><r><t>41_</t></r>" : $"<r><t>{payloads[index]}</t></r>";
                rows.Append($"<row r=\"{row}\"><c r=\"A{row}\" t=\"s\"><v>{(index == 3 ? 0 : index)}</v></c><c r=\"B{row}\" t=\"inlineStr\"><is><t>{payloads[index]}</t></is></c><c r=\"C{row}\" t=\"str\"><v>{payloads[index]}</v></c><c r=\"D{row}\" t=\"inlineStr\"><is>{rich}</is></c></row>");
            }
            Replace(archive, "xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{ns}\"><dimension ref=\"A1:D5\"/><sheetData>{rows}</sheetData></worksheet>", utf16);
            Replace(archive, "xl/sharedStrings.xml", $"<sst xmlns=\"{ns}\" count=\"4\" uniqueCount=\"3\">" + string.Concat(payloads.Take(3).Select(value => $"<si><t>{value}</t></si>")) + "</sst>", utf16);
        }
        byte[] bytes = source.ToArray();
        using (var loaded = ExcelDocument.Load(new MemoryStream(bytes))) {
            for (int row = 0; row < expected.Length; row++)
                for (int column = 1; column <= 4; column++)
                    Assert.Equal(expected[row], loaded.Sheets[0].Cell(row + 2, column).GetValue<string>());
        }
        using (var reader = ExcelDocument.OpenReader(new MemoryStream(bytes))) {
            for (int row = 0; row < expected.Length; row++) {
                Assert.True(reader.Read());
                for (int column = 0; column < 4; column++) {
                    Assert.Equal(expected[row], reader.GetString(column));
#if NET8_0_OR_GREATER
                    byte[] buffer = new byte[128];
                    int count = (int)reader.GetUtf8Bytes(column, 0, buffer, 0, buffer.Length);
                    Assert.Equal(expected[row], Encoding.UTF8.GetString(buffer, 0, count));
#endif
                }
            }
            Assert.False(reader.Read());
        }
        using var owner = ExcelDocumentReader.Open(new MemoryStream(bytes));
        var values = owner.GetSheet("Text").ReadRange("A2:D5");
        for (int row = 0; row < expected.Length; row++)
            for (int column = 0; column < 4; column++) Assert.Equal(expected[row], values[row, column]);
        var records = owner.GetSheet("Text").ReadObjects<LiteralRecord>("A1:D5").ToArray();
        Assert.Equal(expected, records.Select(item => item.Shared));
        Assert.Equal(expected, records.Select(item => item.Inline));
        Assert.Equal(expected, records.Select(item => item.Plain));
        Assert.Equal(expected, records.Select(item => item.Rich));
    }

    [Fact]
    public void TextPastSharedStringPromotionLimitIsEncodedAndReadAsLiteral() {
        using var document = ExcelDocument.Create(new MemoryStream());
        var sheet = document.AddSheet("Text");
        for (int row = 1; row <= 4096; row++) sheet.Cell(row, 1).SetValue("Unique" + row.ToString(CultureInfo.InvariantCulture));
        sheet.Cell(4097, 1).SetValue("_x0041_");
        Assert.Equal("_x0041_", sheet.Cell(4097, 1).GetValue<string>());
        using var source = new MemoryStream();
        document.Save(source, new ExcelSaveOptions { DisableFastPackageWriter = true });
        using (var package = SpreadsheetDocument.Open(new MemoryStream(source.ToArray()), false)) {
            var cell = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single(item => item.CellReference!.Value == "A4097");
            Assert.Equal(CellValues.String, cell.DataType!.Value);
            Assert.Equal("_x005F_x0041_", cell.CellValue!.Text);
        }
        using var loaded = ExcelDocument.Load(new MemoryStream(source.ToArray()));
        Assert.Equal("_x0041_", loaded.Sheets[0].Cell(4097, 1).GetValue<string>());
        loaded.Sheets[0].Cell(4097, 1).SetValue("_x005F_");
        Assert.Equal("_x005F_", loaded.Sheets[0].Cell(4097, 1).GetValue<string>());
        using var resaved = ExcelDocument.Load(new MemoryStream(loaded.ToBytes()));
        Assert.Equal("_x005F_", resaved.Sheets[0].Cell(4097, 1).GetValue<string>());
    }

    [Fact]
    public void RichTextCaseTransformsSemanticTextAndReencodesTheResult() {
        using var document = ExcelDocument.Create();
        var cell = document.AddSheet("Text").Cell(1, 1);
        cell.SetRichText(new ExcelRichTextRun("_x0041_ word"));
        cell.TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.Uppercase, CultureInfo.InvariantCulture);
        Assert.Equal("_X0041_ WORD", cell.GetValue<string>());
        Assert.Equal("_X0041_ WORD", Assert.Single(cell.GetRichText()).Text);
        using var loaded = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal("_X0041_ WORD", loaded.Sheets[0].Cell(1, 1).GetValue<string>());
    }

    [Theory]
    [InlineData("word", "new", "new _x0041_", "new _x0041_")]
    [InlineData("_x0041_", "changed", "word changed", "word changed")]
    public void ReplaceAllMatchesSemanticTextAndPreservesEscapedUnitsAcrossStorageKinds(
        string oldText, string newText, string expectedPrefix, string expectedSplit) {
        const string suffix = " \u0001 \uD800 \uFFFF";
        using var document = ExcelDocument.Load(new MemoryStream(CreateNativeStringFixture("word _x005F_x0041_ _x0001_ _xD800_ _xFFFF_")));
        var sheet = document.Sheets[0];
        for (int column = 1; column <= 5; column++)
            Assert.Equal("word _x0041_" + suffix, sheet.Cell(1, column).GetValue<string>());
        Assert.Equal("word _x0041_", sheet.Cell(1, 6).GetValue<string>());

        Assert.Equal(6, sheet.ReplaceAll(oldText, newText));

        using var reopened = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        for (int column = 1; column <= 5; column++) {
            Assert.Equal(expectedPrefix + suffix, sheet.Cell(1, column).GetValue<string>());
            Assert.Equal(expectedPrefix + suffix, reopened.Sheets[0].Cell(1, column).GetValue<string>());
        }
        Assert.Equal(expectedSplit, reopened.Sheets[0].Cell(1, 6).GetValue<string>());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImportedEscapedUnitsSurviveRichCasingAndSemanticRewrites(bool standardWriter) {
        const string payload = "_x0000__x0001__x000B_ _xD800_X_xDC00_ _xFFFE__xFFFF_ _xD83D__xDE00_ word";
        const string expected = "\0\u0001\u000B \uD800X\uDC00 \uFFFE\uFFFF 😀 WORD";
        const string encoded = "_x0000__x0001__x000B_ _xD800_X_xDC00_ _xFFFE__xFFFF_ 😀 WORD";
        using var document = ExcelDocument.Load(new MemoryStream(CreateNativeStringFixture(payload)));
        var sheet = document.Sheets[0];
        foreach (int column in new[] { 4, 5 }) {
            var cell = sheet.Cell(1, column);
            cell.TransformTextCase(OfficeIMO.Drawing.OfficeTextCase.Uppercase, CultureInfo.InvariantCulture);
            Assert.Equal(expected, cell.GetValue<string>());
            ExcelRichTextRun run = Assert.Single(cell.GetRichText());
            Assert.Equal(expected, run.Text);
            Assert.True(run.Bold);
            sheet.Cell(2, column).SetRichText(run);
            sheet.Cell(3, column).SetValue(cell.GetValue<string>());
        }

        using var saved = new MemoryStream();
        document.Save(saved, new ExcelSaveOptions { DisableFastPackageWriter = standardWriter });
        using (var package = SpreadsheetDocument.Open(new MemoryStream(saved.ToArray()), false)) {
            var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>()
                .Where(cell => cell.CellReference?.Value == "D1" || cell.CellReference?.Value == "E1").ToArray();
            Assert.Equal(2, cells.Length);
            foreach (Cell cell in cells) Assert.Equal(encoded, Assert.Single(cell.InlineString!.Elements<Run>()).Text!.Text);
        }
        using var reopened = ExcelDocument.Load(new MemoryStream(saved.ToArray()));
        for (int row = 1; row <= 3; row++)
            foreach (int column in new[] { 4, 5 })
                Assert.Equal(expected, reopened.Sheets[0].Cell(row, column).GetValue<string>());
    }

    [Fact]
    public void TemplateBindingPreservesUnrelatedImportedEscapedUnits() {
        using var document = ExcelDocument.Load(new MemoryStream(CreateNativeStringFixture("{{Name}} _x0001_ _xD800_ _x005F_x0041_")));
        var cells = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToArray();
        string[] before = cells.Select(cell => cell.OuterXml).ToArray();
        Assert.Throws<ArgumentException>(() => document.Sheets[0].ApplyTemplate(
            new Dictionary<string, object?> { ["Name"] = new string('A', 32767) }));
        Assert.Equal(before, cells.Select(cell => cell.OuterXml).ToArray());
        Assert.Equal(5, document.Sheets[0].ApplyTemplate(new Dictionary<string, object?> { ["Name"] = "Ada" }));
        using var reopened = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        for (int column = 1; column <= 5; column++)
            Assert.Equal("Ada \u0001 \uD800 _x0041_", reopened.Sheets[0].Cell(1, column).GetValue<string>());
    }

    // Independent native payloads exercise semantic consumers without using the encoder under test.
    internal static byte[] CreateNativeStringFixture(string payload) {
        using var document = ExcelDocument.Create();
        document.AddSheet("Text").Cell(1, 1).SetValue("seed");
        using var source = new MemoryStream();
        document.Save(source);
        const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        string rich = "<r><rPr><b/></rPr><t>" + payload + "</t></r>";
        using (var archive = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true)) {
            Replace(archive, "xl/sharedStrings.xml", $"<sst xmlns=\"{ns}\" count=\"2\" uniqueCount=\"2\"><si><t>{payload}</t></si><si>{rich}</si></sst>", false);
            Replace(archive, "xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{ns}\"><dimension ref=\"A1:F1\"/><sheetData><row r=\"1\">"
                + "<c r=\"A1\" t=\"s\"><v>0</v></c>"
                + $"<c r=\"B1\" t=\"inlineStr\"><is><t>{payload}</t></is></c>"
                + $"<c r=\"C1\" t=\"str\"><v>{payload}</v></c>"
                + $"<c r=\"D1\" t=\"inlineStr\"><is>{rich}</is></c>"
                + "<c r=\"E1\" t=\"s\"><v>1</v></c>"
                + "<c r=\"F1\" t=\"inlineStr\"><is><r><t>word _x00</t></r><r><t>41_</t></r></is></c>"
                + "</row></sheetData></worksheet>", false);
        }
        return source.ToArray();
    }

    public sealed class LiteralRecord {
        public string Shared { get; set; } = string.Empty;
        public string Inline { get; set; } = string.Empty;
        public string Plain { get; set; } = string.Empty;
        public string Rich { get; set; } = string.Empty;
    }

    private static void Replace(ZipArchive archive, string path, string xml, bool utf16) {
        archive.GetEntry(path)?.Delete();
        using var stream = archive.CreateEntry(path).Open();
        using var writer = new StreamWriter(stream, utf16 ? Encoding.Unicode : new UTF8Encoding(false));
        writer.Write(xml);
    }
}
