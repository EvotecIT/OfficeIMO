using System.IO.Compression;
using System.Xml;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    private const int IndexedSharedStringFillerCount = 512;

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void OpenDataReader_LargePlainSharedStringsPreserveFieldsAndProjectionFallbacks(
        bool stored, bool utf16Worksheet) {
        string path = CreateIndexedSharedStringWorkbook(
            "<si><t xml:space=\"preserve\"> last </t></si>",
            includeLastRow: true, stored: stored, utf16Worksheet: utf16Worksheet);
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.Equal("Header", reader.GetName(0));
            Assert.True(reader.Read());
            Assert.Equal("Ready", reader.GetString(0));
#if NET8_0_OR_GREATER
            if (!utf16Worksheet) {
                Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
                Assert.Equal("Ready", Encoding.UTF8.GetString(text));
            }
#endif
            Assert.True(reader.Read());
            Assert.Equal(" last ", reader.GetString(0));
#if NET8_0_OR_GREATER
            if (!utf16Worksheet) {
                Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
                Assert.Equal(" last ", Encoding.UTF8.GetString(text));
            }
#endif
            Assert.False(reader.Read());
            reader.Dispose();
            File.Delete(path);
            Assert.False(File.Exists(path));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("<si><t>Żółw 🐢</t></si>", "Żółw 🐢")]
    [InlineData("<si><t>A&amp;B &#xD;tail</t></si>", "A&B \rtail")]
    [InlineData("<si><r><t>first</t></r><rPh sb=\"0\" eb=\"1\"><t>phonetic</t></rPh><r><t>second</t></r></si>", "firstsecond")]
    [InlineData("<si><t>A\r\nB</t></si>", "A\nB")]
    [InlineData("<si><t> </t></si>", "")]
    [InlineData("<si><t>_x0041_</t></si>", "_x0041_")]
    public void OpenDataReader_LargeSharedTableWithLateComplexTextKeepsCanonicalFallback(
        string lastItem, string expected) {
        string path = CreateIndexedSharedStringWorkbook(lastItem, includeLastRow: true);
        try {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(path);
            Assert.True(reader.Read());
            Assert.Equal("Ready", reader.GetString(0));
            Assert.True(reader.Read());
            Assert.Equal(expected, reader.GetString(0));
#if NET8_0_OR_GREATER
            Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
            Assert.Equal(expected, Encoding.UTF8.GetString(text));
#endif
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("<si><t>unclosed</si>")]
    [InlineData("<si><t>invalid &unknown;</t></si>")]
    [InlineData("<si><t>invalid\u0001</t></si>")]
    [InlineData("<si><t>invalid]]></t></si>")]
    public void OpenDataReader_LargeSharedTableRejectsMalformedUnreferencedTailBeforeDelivery(string lastItem) {
        string path = CreateIndexedSharedStringWorkbook(lastItem);
        try {
            Assert.Throws<XmlException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("Items")]
    [InlineData("ItemCharacters")]
    [InlineData("AggregateCharacters")]
    public void OpenDataReader_LargeSharedTableEnforcesUnreferencedEntryLimits(string limit) {
        string path = CreateIndexedSharedStringWorkbook("<si><t>unreferenced over budget</t></si>");
        try {
            var options = new ExcelReadOptions {
                MaxSharedStringItems = limit == "Items" ? IndexedSharedStringFillerCount + 2 : IndexedSharedStringFillerCount + 3,
                MaxSharedStringItemCharacters = limit == "ItemCharacters" ? 20 : 100,
                MaxSharedStringCharacters = limit == "AggregateCharacters" ? IndexedSharedStringFillerCount * 16L + 11 : 100_000
            };
            Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(path, options));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OpenDataReader_LargeSharedTableBoundsIndexesByActualEntries() {
        string path = CreateIndexedSharedStringWorkbook("<si><t>last</t></si>",
            includeLastRow: true, lastIndex: IndexedSharedStringFillerCount + 3,
            declaredCount: "2147483647");
        try {
            Assert.Throws<InvalidDataException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    private static string CreateIndexedSharedStringWorkbook(
        string lastItem,
        bool includeLastRow = false,
        bool stored = false,
        bool utf16Worksheet = false,
        int lastIndex = IndexedSharedStringFillerCount + 2,
        string? declaredCount = null,
        int fillerCount = IndexedSharedStringFillerCount) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.IndexedSst.{Guid.NewGuid():N}.xlsx");
        using (var document = ExcelDocument.Create(path)) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Header");
            sheet.CellValue(2, 1, "Ready");
            document.Save();
        }
        string count = declaredCount ?? (fillerCount + 3).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var sharedStrings = new StringBuilder(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"" + count +
            "\" uniqueCount=\"" + count + "\"><si><t>Header</t></si><si><t>Ready</t></si>");
        for (int index = 0; index < fillerCount; index++) {
            sharedStrings.Append("<si><t>unreferenced-").Append(index.ToString("D3", System.Globalization.CultureInfo.InvariantCulture)).Append("</t></si>");
        }
        sharedStrings.Append(lastItem).Append("</sst>");
        string rows = "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>" +
            "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>1</v></c></row>" +
            (includeLastRow ? "<row r=\"3\"><c r=\"A3\" t=\"s\"><v>" + lastIndex + "</v></c></row>" : string.Empty);
        string worksheet = "<?xml version=\"1.0\" encoding=\"" + (utf16Worksheet ? "utf-16" : "UTF-8") + "\"?>" +
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>" + rows + "</sheetData></worksheet>";
        using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update)) {
            ReplaceIndexedSharedStringPart(archive, "xl/sharedStrings.xml", Encoding.UTF8.GetBytes(sharedStrings.ToString()), stored);
            byte[] worksheetBytes = utf16Worksheet
                ? Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(worksheet)).ToArray()
                : Encoding.UTF8.GetBytes(worksheet);
            ReplaceIndexedSharedStringPart(archive, "xl/worksheets/sheet1.xml", worksheetBytes, stored);
        }
        return path;
    }

    private static void ReplaceIndexedSharedStringPart(ZipArchive archive, string name, byte[] bytes, bool stored) {
        archive.GetEntry(name)!.Delete();
        using Stream output = archive.CreateEntry(name, stored ? CompressionLevel.NoCompression : CompressionLevel.Optimal).Open();
        output.Write(bytes, 0, bytes.Length);
    }
}
