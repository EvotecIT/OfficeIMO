using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfWordTableImportDefaultsTests {
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Theory]
    [InlineData(0)]
    [InlineData(300)]
    public void Rtf_To_Word_Import_Persists_Modern_Gap_Profile(int gapTwips) {
        RtfDocument source = RtfDocument.Load(GapFixturePath(gapTwips, ".rtf"));
        RtfTable sourceTable = Assert.Single(source.Blocks.OfType<RtfTable>());
        Assert.Equal(gapTwips, Assert.Single(sourceTable.Rows).CellGapTwips);
        Assert.Equal(0, sourceTable.Rows[0].LeftIndentTwips);
        Assert.Null(sourceTable.Rows[0].Cells[0].PaddingLeftTwips);

        using WordDocument mapped = source.ToWordDocument();
        XDocument document = ReadPart(mapped.ToBytes(), "word/document.xml");
        XElement table = Assert.Single(document.Descendants(W + "tbl"));

        AssertNeutralTable(table);
        Assert.Equal(WordCompatibilityMode.Word2013, mapped.CompatibilitySettings.CompatibilityMode);
        AssertGapProperties(table, gapTwips, expectedIndentTwips: 0);
        XElement cell = Assert.Single(table.Descendants(W + "tc"));
        Assert.Null(cell.Element(W + "tcPr")?.Element(W + "tcMar")?.Element(W + "left"));
        Assert.Null(cell.Element(W + "tcPr")?.Element(W + "tcMar")?.Element(W + "right"));
        AssertEffectiveMargins(table, cell, 0, gapTwips, 0, gapTwips);
        Assert.Empty(table.Descendants(W + "tcBorders"));
        Assert.Equal(ExpectedCellText, WithoutWhiteSpace(string.Concat(table.Descendants(W + "t").Select(item => item.Value))));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(300)]
    public void Word_Rtf_Import_Recovers_Native_Gap_Profile(int gapTwips) {
        byte[] fixture = File.ReadAllBytes(GapFixturePath(gapTwips, ".native.docx"));
        XElement nativeTable = Assert.Single(ReadPart(fixture, "word/document.xml").Descendants(W + "tbl"));
        AssertGapProperties(nativeTable, gapTwips);
        using var fixtureStream = new MemoryStream(fixture, writable: false);
        using WordDocument native = WordDocument.Load(fixtureStream);
        Assert.Equal(WordCompatibilityMode.Word2003, native.CompatibilitySettings.CompatibilityMode);

        RtfDocument converted = native.ToRtfDocument();
        IEnumerable<IRtfBlock> blocks = converted.Sections.Count > 0
            ? converted.Sections.SelectMany(section => section.Blocks)
            : converted.Blocks;
        RtfTable table = Assert.Single(blocks.OfType<RtfTable>());
        RtfTableRow row = Assert.Single(table.Rows);
        Assert.Equal(gapTwips, row.CellGapTwips);
        Assert.Equal(0, row.LeftIndentTwips);
        RtfTableCell cell = Assert.Single(row.Cells);
        Assert.Null(cell.PaddingLeftTwips);
        Assert.Null(cell.PaddingRightTwips);
        Assert.Equal(ExpectedCellText, WithoutWhiteSpace(string.Concat(cell.Paragraphs.Select(paragraph => paragraph.ToPlainText()))));

        using WordDocument returned = converted.ToWordDocument();
        XElement returnedTable = Assert.Single(ReadPart(returned.ToBytes(), "word/document.xml").Descendants(W + "tbl"));
        AssertGapProperties(returnedTable, gapTwips, expectedIndentTwips: 0);
        Assert.Equal(ExpectedCellText, WithoutWhiteSpace(string.Concat(returnedTable.Descendants(W + "t").Select(item => item.Value))));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Rtf_To_Word_Import_Applies_Defaults_In_Body_Section_And_Nested_Tables(bool section, bool nested) {
        RtfDocument source = RtfDocument.Create();
        RtfTable outer = section ? source.AddSection().AddTable(1, 1) : source.AddTable(1, 1);
        outer.Rows[0].SetCellGap(0).SetLeftIndent(0);
        RtfTable inner = nested ? outer.Rows[0].Cells[0].AddTable(1, 1) : outer;
        inner.Rows[0].SetCellGap(0).SetLeftIndent(0);
        RtfParagraph paragraph = inner.Rows[0].Cells[0].AddParagraph("Cell");
        RtfNote footnote = source.AddNote(RtfNoteKind.Footnote);
        footnote.AddParagraph("Cell footnote");
        paragraph.AddNoteReference(footnote, "1");
        RtfNote endnote = source.AddNote(RtfNoteKind.Endnote);
        endnote.AddParagraph("Cell endnote");
        paragraph.AddNoteReference(endnote, "2");

        using WordDocument mapped = source.ToWordDocument();
        byte[] docx = mapped.ToBytes();
        XDocument document = ReadPart(docx, "word/document.xml");
        XElement[] tables = document.Descendants(W + "tbl").ToArray();

        Assert.Equal(nested ? 2 : 1, tables.Length);
        foreach (XElement table in tables) {
            AssertNeutralTable(table);
            foreach (XElement cell in table.Elements(W + "tr").SelectMany(row => row.Elements(W + "tc"))) {
                AssertEffectiveMargins(table, cell, 0, 0, 0, 0);
            }
        }
        // Notes referenced inside cells still use the existing paragraph-only note mapper.
        Assert.Contains("Cell footnote", ReadPart(docx, "word/footnotes.xml").ToString());
        Assert.Contains("Cell endnote", ReadPart(docx, "word/endnotes.xml").ToString());
        Assert.Single(document.Descendants(W + "footnoteReference"));
        Assert.Single(document.Descendants(W + "endnoteReference"));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(300)]
    public void Rtf_To_Word_Import_Preserves_Explicit_Cell_Padding_And_Borders(int gapTwips) {
        string rtf = @"{\rtf1\ansi{\colortbl;\red68\green114\blue196;}\trowd\trgaph" +
            gapTwips.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            @"\trleft0\clpadt120\clpadft3\clpadl180\clpadfl3\clpadb240\clpadfb3\clpadr300\clpadfr3\clbrdrt\brdrs\brdrw12\brdrcf1\cellx2400\cellx4800\clpadl0\clpadfl3\clpadr0\clpadfr3\cellx7200\pard\intbl Explicit\cell\pard\intbl Plain\cell\pard\intbl Zero\cell\row}";
        using WordDocument mapped = RtfDocument.Read(rtf).Document.ToWordDocument();
        XElement table = Assert.Single(ReadPart(mapped.ToBytes(), "word/document.xml").Descendants(W + "tbl"));
        XElement[] cells = table.Descendants(W + "tc").ToArray();

        AssertNeutralTable(table);
        Assert.Equal(3, cells.Length);
        AssertMargins(cells[0], 120, 180, 240, 300);
        AssertEffectiveMargins(table, cells[1], 0, gapTwips, 0, gapTwips);
        AssertMargins(cells[2], 0, 0, 0, 0);
        XElement top = cells[0].Element(W + "tcPr")!.Element(W + "tcBorders")!.Element(W + "top")!;
        Assert.Equal("single", (string?)top.Attribute(W + "val"));
        Assert.Equal("12", (string?)top.Attribute(W + "sz"));
        Assert.Equal("4472C4", (string?)top.Attribute(W + "color"));
        Assert.Null(cells[1].Element(W + "tcPr")?.Element(W + "tcBorders"));
    }

    [Fact]
    public void Rtf_To_Word_Import_Preserves_Row_Padding_With_Cell_Overrides() {
        RtfDocument source = RtfDocument.Create();
        RtfTableRow row = source.AddTable(1, 3).Rows[0];
        row.SetCellGap(300).SetLeftIndent(0).SetPadding(topTwips: 60, leftTwips: 90, bottomTwips: 80, rightTwips: 110);
        row.Cells[0].SetPadding(leftTwips: 180, rightTwips: 240);
        row.Cells[2].SetPadding(topTwips: 0, leftTwips: 0, bottomTwips: 0, rightTwips: 0);

        using WordDocument mapped = source.ToWordDocument();
        XElement table = Assert.Single(ReadPart(mapped.ToBytes(), "word/document.xml").Descendants(W + "tbl"));
        XElement[] cells = table.Elements(W + "tr").SelectMany(item => item.Elements(W + "tc")).ToArray();
        AssertGapProperties(table, 300, expectedIndentTwips: 0);
        AssertEffectiveMargins(table, cells[0], 60, 180, 80, 240);
        AssertEffectiveMargins(table, cells[1], 60, 90, 80, 110);
        AssertEffectiveMargins(table, cells[2], 0, 0, 0, 0);
    }

    [Theory]
    [InlineData(0, 300)]
    [InlineData(0, null)]
    [InlineData(null, null)]
    public void Rtf_To_Word_Import_Reports_Mixed_Or_Missing_Row_Gaps(int? firstGap, int? secondGap) {
        RtfDocument source = RtfDocument.Create();
        RtfTable table = source.AddTable(2, 1);
        table.Rows[0].CellGapTwips = firstGap;
        table.Rows[1].CellGapTwips = secondGap;
        table.Rows[0].Cells[0].AddParagraph("FIRST");
        table.Rows[1].Cells[0].AddParagraph("SECOND");

        RtfConversionResult<WordDocument> result = source.ToWordDocumentResult();
        using WordDocument mapped = result.Value;
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "RtfWordTableCellGapFlattened");
        Assert.True(result.Report.HasLoss);
        Assert.Equal(firstGap, table.Rows[0].CellGapTwips);
        Assert.Equal(secondGap, table.Rows[1].CellGapTwips);
        string text = string.Concat(ReadPart(mapped.ToBytes(), "word/document.xml").Descendants(W + "t").Select(item => item.Value));
        Assert.Equal("FIRSTSECOND", text);
    }

    [Fact]
    public void Rtf_Import_Defaults_Do_Not_Change_Word_Authoring_Or_Caller_Style() {
        RtfDocument source = RtfDocument.Create();
        source.AddTable(1, 1).Rows[0].SetCellGap(0);
        using WordDocument mapped = source.ToWordDocument();
        mapped.AddTable(1, 1); // Existing public Word authoring default.
        WordTable callerStyled = mapped.AddTable(1, 1, WordTableStyle.TableGrid);
        callerStyled.Rows[0].Cells[0].MarginLeftWidth = 180;
        callerStyled.Rows[0].Cells[0].Borders.TopStyle = WordBorderStyle.Double;
        XElement[] tables = ReadPart(mapped.ToBytes(), "word/document.xml").Descendants(W + "tbl").ToArray();

        Assert.Equal(3, tables.Length);
        AssertNeutralTable(tables[0]);
        Assert.Equal("TableGrid", (string?)tables[1].Element(W + "tblPr")?.Element(W + "tblStyle")?.Attribute(W + "val"));
        Assert.Equal("TableGrid", (string?)tables[2].Element(W + "tblPr")?.Element(W + "tblStyle")?.Attribute(W + "val"));
        Assert.Null(tables[1].Element(W + "tblPr")?.Element(W + "tblBorders"));
        XElement callerCellProperties = Assert.Single(tables[2].Descendants(W + "tcPr"));
        Assert.Equal("180", (string?)callerCellProperties.Element(W + "tcMar")?.Element(W + "left")?.Attribute(W + "w"));
        Assert.Equal("double", (string?)callerCellProperties.Element(W + "tcBorders")?.Element(W + "top")?.Attribute(W + "val"));
    }

    private const string ExpectedCellText = "TABLE-MIN-ATABLE-MIN-TALLEND-TABLE-MINTABLE-MIN-BTABLE-EXACT-ATABLE-EXACT-TALLEND-TABLE-EXACTTABLE-EXACT-B";

    private static string GapFixturePath(int gapTwips, string extension) => Path.Combine(
        AppContext.BaseDirectory, "Documents", "RtfTableGapNative",
        "table-cell-minimum-exact-gap" + gapTwips.ToString(System.Globalization.CultureInfo.InvariantCulture) + extension);

    private static string WithoutWhiteSpace(string value) => new string(value.Where(character => !char.IsWhiteSpace(character)).ToArray());

    private static void AssertGapProperties(XElement table, int gapTwips, int? expectedIndentTwips = null) {
        XElement properties = table.Element(W + "tblPr")!;
        Assert.NotNull(properties);
        Assert.Null(properties.Element(W + "tblCellSpacing"));
        XElement margins = properties.Element(W + "tblCellMar")!;
        Assert.NotNull(margins);
        AssertMargin(margins, "left", gapTwips);
        AssertMargin(margins, "right", gapTwips);
        XElement? indentation = properties.Element(W + "tblInd");
        int expectedIndent = expectedIndentTwips ?? gapTwips;
        if (expectedIndent == 0) {
            Assert.Null(indentation);
        } else {
            Assert.NotNull(indentation);
            Assert.Equal(expectedIndent.ToString(System.Globalization.CultureInfo.InvariantCulture), (string?)indentation!.Attribute(W + "w"));
            Assert.Equal("dxa", (string?)indentation.Attribute(W + "type"));
        }
    }

    private static void AssertEffectiveMargins(XElement table, XElement cell, int top, int left, int bottom, int right) {
        foreach (var expected in new[] { (Side: "top", Value: top), (Side: "left", Value: left),
            (Side: "bottom", Value: bottom), (Side: "right", Value: right) }) {
            XElement? margin = cell.Element(W + "tcPr")?.Element(W + "tcMar")?.Element(W + expected.Side) ??
                table.Element(W + "tblPr")?.Element(W + "tblCellMar")?.Element(W + expected.Side);
            Assert.NotNull(margin);
            Assert.Equal(expected.Value.ToString(System.Globalization.CultureInfo.InvariantCulture), (string?)margin!.Attribute(W + "w"));
            Assert.Equal("dxa", (string?)margin.Attribute(W + "type"));
        }
    }

    private static void AssertNeutralTable(XElement table) {
        XElement properties = table.Element(W + "tblPr")!;
        XElement borders = properties.Element(W + "tblBorders")!;
        Assert.NotNull(borders);
        Assert.Equal(6, borders.Elements().Count());
        foreach (string side in new[] { "top", "bottom", "left", "right", "insideH", "insideV" }) {
            Assert.Equal("nil", (string?)borders.Element(W + side)?.Attribute(W + "val"));
        }
    }

    private static void AssertMargins(XElement cell, int top, int left, int bottom, int right) {
        XElement margins = cell.Element(W + "tcPr")!.Element(W + "tcMar")!;
        Assert.NotNull(margins);
        AssertMargin(margins, "top", top);
        AssertMargin(margins, "left", left);
        AssertMargin(margins, "bottom", bottom);
        AssertMargin(margins, "right", right);
    }

    private static void AssertMargin(XElement margins, string side, int expected) {
        XElement margin = margins.Element(W + side)!;
        Assert.NotNull(margin);
        Assert.Equal(expected.ToString(System.Globalization.CultureInfo.InvariantCulture), (string?)margin.Attribute(W + "w"));
        Assert.Equal("dxa", (string?)margin.Attribute(W + "type"));
    }

    private static XDocument ReadPart(byte[] docx, string name) {
        using var bytes = new MemoryStream(docx);
        using var archive = new ZipArchive(bytes, ZipArchiveMode.Read);
        ZipArchiveEntry entry = archive.GetEntry(name)!;
        Assert.NotNull(entry);
        using Stream part = entry.Open();
        return XDocument.Load(part);
    }
}
