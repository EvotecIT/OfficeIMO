using System.Data;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Excel.Tests;

/// <summary>
/// Hand-written package and XML scanners must agree with the OPC and XML rules: part names compare
/// case-insensitively, and comments or processing instructions may precede or start any part's root.
/// </summary>
public sealed class ExcelPartMarkupAndNameVariantTests {
    private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    // Tag-like comment and processing-instruction text that a byte scanner could mistake for real markup.
    private const string DecoyMarkup =
        "<!-- <worksheet><sheetData><row r=\"9\"><c r=\"A9\"><v>9</v></c></row></sheetData></worksheet>"
        + " <x:sheetData><x:row r=\"9\"><x:c r=\"A9\"><x:v>9</x:v></x:c></x:row></x:sheetData>"
        + " <sst><si><t>Decoy</t></si></sst> <x:si><x:t>Decoy</x:t></x:si> <x:table name=\"Decoy\"> -->"
        + "<?decoy <x:worksheet> <sst> <table> ?>\r\n";

    [Theory]
    [InlineData("x:", true, false, "xl/worksheets/sheet1.xml", "xl/tables/table1.xml")]
    [InlineData("x:", false, true, "xl/worksheets/sheet1.xml", "xl/tables/table1.xml")]
    [InlineData("", true, true, "xl/worksheets/sheet1.xml", "xl/tables/table1.xml")]
    [InlineData("", false, false, "xl/Worksheets/Sheet1.xml", "xl/Tables/Table1.xml")]
    [InlineData("x:", true, false, "xl/Worksheets/Sheet1.xml", "xl/Tables/Table1.xml")]
    public void ReadersAgreeOnPrefixedPartsDecoyMarkupAndPartNameCase(
        string prefix, bool decoyBeforeRoot, bool decoyInsideRoot, string worksheetPart, string tablePart) {
        byte[] package = CreatePackage(prefix, decoyBeforeRoot ? DecoyMarkup : "", decoyInsideRoot ? DecoyMarkup : "",
            worksheetPart, tablePart);
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.Excel.PartVariants.{Guid.NewGuid():N}.xlsx");
        File.WriteAllBytes(path, package);
        try {
            using (ExcelDocument document = ExcelDocument.Load(path, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                ExcelTableInfo table = Assert.Single(document.GetTables());
                Assert.Equal("Items", table.Name);
                ExcelSheet sheet = document.GetSheet("Data");
                Assert.True(sheet.TryGetCellText(2, 1, out string text));
                Assert.Equal("Alpha", text);
                DataTable rows = sheet.TableToDataTable("Items");
                Assert.Equal(new object[] { "Beta", 2d }, new[] { rows.Rows[1]["Name"], Convert.ToDouble(rows.Rows[1]["Amount"]) });
            }

            using (ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(path)) {
                Assert.True(reader.Read());
                Assert.Equal("Alpha", reader.GetString(0));
                Assert.Equal(1.5d, Convert.ToDouble(reader.GetValue(1)));
                Assert.Equal(DateTime.FromOADate(45000), reader.GetValue(2));
                Assert.True(reader.Read());
                Assert.Equal("Beta", reader.GetString(0));
                Assert.False(reader.Read());
            }

            using (ExcelDocumentReader reader = ExcelDocumentReader.Open(package)) {
                DataTable table = reader.ReadTableAsDataTable("Items");
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal("Alpha", table.Rows[0]["Name"]);
                Assert.Equal(DateTime.FromOADate(45001), table.Rows[1]["When"]);
                object?[]? lastRow = reader.GetSheet("Data").ReadRows("A1:C3").Last();
                Assert.Equal("Beta", lastRow![0]);
            }
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void TableCatalogIgnoresMetadataInExtensionsAndUnprocessedAlternateContent() {
        const string alternate = """
            <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:future="urn:future">
              <mc:Choice Requires="future"><tableStyleInfo name="DecoyChoice" /></mc:Choice>
              <mc:Fallback><tableStyleInfo name="DecoyFallback" /></mc:Fallback>
            </mc:AlternateContent>
            """;
        const string extension = """
            <extLst><ext uri="{CC7BAA82-25DE-40CE-8FA7-82D7F8675300}">
              <tableColumns count="1"><tableColumn id="99" name="Decoy" /></tableColumns>
              <tableStyleInfo name="DecoyExtension" />
            </ext></extLst>
            """;
        const string totalsAttributes = """
             xmlns:future="urn:future" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
             mc:Ignorable="future" future:totalsRowFunction="sum" totalsRowFunction="max"
            """;
        byte[] package = CreatePackage("", "", "", "xl/worksheets/sheet1.xml", "xl/tables/table1.xml",
            alternate, extension, totalsAttributes);

        using ExcelDocumentReader reader = ExcelDocumentReader.Open(package);
        ExcelTableInfo info = Assert.Single(reader.GetTables());
        Assert.Equal("TableStyleMedium2", info.StyleName);
        Assert.Equal(new[] { "Name", "Amount", "When" }, info.Columns.Select(column => column.Name));
        Assert.Equal("max", info.Columns[0].TotalsRowFunction);
        Assert.True(info.HasAutoFilter);
        Assert.Equal("Alpha", reader.ReadTableAsDataTable("Items").Rows[0]["Name"]);
    }

    [Theory]
    [InlineData("", false)]
    [InlineData("<wrong />", true)]
    public void TableCatalogDistinguishesEmptyPartsFromWrongRootElements(string tableMarkup, bool invalid) {
        byte[] package = CreatePackage("", "", "", "xl/worksheets/sheet1.xml", "xl/tables/table1.xml", tableMarkup: tableMarkup);
        using ExcelDocumentReader reader = ExcelDocumentReader.Open(package);
        if (invalid) {
            Assert.Throws<InvalidDataException>(() => reader.GetTables());
        } else {
            Assert.Empty(reader.GetTables());
        }
    }

    private static byte[] CreatePackage(string p, string beforeRoot, string insideRoot, string worksheetPart, string tablePart,
        string tableStart = "", string tableEnd = "", string firstColumnAttributes = "", string? tableMarkup = null) {
        const string declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
        const string relationshipsStart = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
        string ns = p.Length == 0 ? $"xmlns=\"{MainNamespace}\"" : $"xmlns:{p.TrimEnd(':')}=\"{MainNamespace}\"";
        string head = declaration + beforeRoot;
        string worksheetFile = worksheetPart.Substring(worksheetPart.LastIndexOf('/') + 1);
        string worksheetFolder = worksheetPart.Substring(0, worksheetPart.LastIndexOf('/'));
        var parts = new (string Name, string Xml)[] {
            ("[Content_Types].xml", declaration + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                + $"<Override PartName=\"/{worksheetPart}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                + "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
                + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
                + $"<Override PartName=\"/{tablePart}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/></Types>"),
            ("_rels/.rels", declaration + relationshipsStart
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"),
            // Relationships use the canonical lower-case names whatever case the stored parts use.
            ("xl/_rels/workbook.xml.rels", declaration + relationshipsStart
                + $"<Relationship Id=\"rId1\" Type=\"{RelationshipNamespace}/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                + $"<Relationship Id=\"rId2\" Type=\"{RelationshipNamespace}/sharedStrings\" Target=\"sharedStrings.xml\"/>"
                + $"<Relationship Id=\"rId3\" Type=\"{RelationshipNamespace}/styles\" Target=\"styles.xml\"/></Relationships>"),
            (worksheetFolder + "/_rels/" + worksheetFile + ".rels", declaration + relationshipsStart
                + $"<Relationship Id=\"rId1\" Type=\"{RelationshipNamespace}/table\" Target=\"../tables/table1.xml\"/></Relationships>"),
            ("xl/workbook.xml", head + $"<{p}workbook {ns} xmlns:r=\"{RelationshipNamespace}\">{insideRoot}"
                + $"<{p}sheets><{p}sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></{p}sheets></{p}workbook>"),
            (worksheetPart, head + $"<{p}worksheet {ns} xmlns:r=\"{RelationshipNamespace}\">{insideRoot}<{p}dimension ref=\"A1:C3\"/>{insideRoot}<{p}sheetData>"
                + $"<{p}row r=\"1\">{SharedString(p, "A1", 0)}{SharedString(p, "B1", 1)}{SharedString(p, "C1", 4)}</{p}row>"
                + $"<{p}row r=\"2\">{SharedString(p, "A2", 2)}<{p}c r=\"B2\"><{p}v>1.5</{p}v></{p}c><{p}c r=\"C2\" s=\"1\"><{p}v>45000</{p}v></{p}c></{p}row>"
                + $"<{p}row r=\"3\">{SharedString(p, "A3", 3)}<{p}c r=\"B3\"><{p}v>2</{p}v></{p}c><{p}c r=\"C3\" s=\"1\"><{p}v>45001</{p}v></{p}c></{p}row>"
                + $"</{p}sheetData><{p}tableParts count=\"1\"><{p}tablePart r:id=\"rId1\"/></{p}tableParts></{p}worksheet>"),
            ("xl/sharedStrings.xml", head + $"<{p}sst {ns} count=\"5\" uniqueCount=\"5\">{insideRoot}"
                + $"<{p}si><{p}t>Name</{p}t></{p}si><{p}si><{p}t>Amount</{p}t></{p}si>"
                + $"<{p}si><{p}r><{p}t>Al</{p}t></{p}r><{p}r><{p}t>pha</{p}t></{p}r></{p}si>"
                + $"<{p}si><{p}t>Beta</{p}t></{p}si><{p}si><{p}t>When</{p}t></{p}si></{p}sst>"),
            ("xl/styles.xml", head + $"<{p}styleSheet {ns}>{insideRoot}"
                + $"<{p}fonts count=\"1\"><{p}font><{p}sz val=\"11\"/><{p}name val=\"Calibri\"/></{p}font></{p}fonts>"
                + $"<{p}fills count=\"2\"><{p}fill><{p}patternFill patternType=\"none\"/></{p}fill><{p}fill><{p}patternFill patternType=\"gray125\"/></{p}fill></{p}fills>"
                + $"<{p}borders count=\"1\"><{p}border><{p}left/><{p}right/><{p}top/><{p}bottom/><{p}diagonal/></{p}border></{p}borders>"
                + $"<{p}cellStyleXfs count=\"1\"><{p}xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></{p}cellStyleXfs>"
                + $"<{p}cellXfs count=\"2\"><{p}xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
                + $"<{p}xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></{p}cellXfs></{p}styleSheet>"),
            (tablePart, tableMarkup ?? (head + $"<{p}table {ns} id=\"1\" name=\"Items\" displayName=\"Items\" ref=\"A1:C3\">{insideRoot}{tableStart}<{p}autoFilter ref=\"A1:C3\"/>"
                + $"<{p}tableColumns count=\"3\"><{p}tableColumn id=\"1\" name=\"Name\"{firstColumnAttributes}/><{p}tableColumn id=\"2\" name=\"Amount\"/><{p}tableColumn id=\"3\" name=\"When\"/></{p}tableColumns>"
                + $"<{p}tableStyleInfo name=\"TableStyleMedium2\" showRowStripes=\"1\"/>{tableEnd}</{p}table>"))
        };

        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (var (name, xml) in parts) {
                using Stream entry = archive.CreateEntry(name).Open();
                byte[] bytes = Encoding.UTF8.GetBytes(xml);
                entry.Write(bytes, 0, bytes.Length);
            }
        }
        return output.ToArray();
    }

    private static string SharedString(string p, string reference, int index) =>
        $"<{p}c r=\"{reference}\" t=\"s\"><{p}v>{index}</{p}v></{p}c>";
}
