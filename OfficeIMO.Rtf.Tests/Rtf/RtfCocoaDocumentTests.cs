using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

/// <summary>The dialect of TextEdit and textutil: flat font tables, face names and backslash line endings.</summary>
public class RtfCocoaDocumentTests {
    private const string Report =
        "{\\rtf1\\ansi\\ansicpg1250\\cocoartf2907\n" +
        "{\\fonttbl\\f0\\froman\\fcharset0 Times-Bold;\\f1\\froman\\fcharset0 Times-Roman;}\n" +
        "{\\colortbl;\\red255\\green255\\blue255;\\red0\\green0\\blue0;}\n" +
        "{\\*\\expandedcolortbl;;\\cssrgb\\c0\\c0\\c0;}\n" +
        "{\\*\\listtable{\\list\\listtemplateid1\\listhybrid{\\listlevel\\levelnfc23\\levelnfcn23\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat0\\levelspace360\\levelindent0{\\*\\levelmarker \\{disc\\}}{\\leveltext\\leveltemplateid1\\'01\\uc0\\u8226 ;}{\\levelnumbers;}\\fi-360\\li720\\lin720 }{\\listname ;}\\listid1}}\n" +
        "{\\*\\listoverridetable{\\listoverride\\listid1\\listoverridecount0\\ls1}}\n" +
        "\\deftab720\n" +
        "\\pard\\pardeftab720\\sa321\\partightenfactor0\n\n" +
        "\\f0\\b\\fs48 \\cf0 Heading\\\n" +
        "\\pard\\pardeftab720\\sa240\\partightenfactor0\n\n" +
        "\\f1\\b0\\fs24 \\cf0 Body text\\\n" +
        "\\pard\\tx220\\tx720\\pardeftab720\\li720\\fi-720\\partightenfactor0\n" +
        "\\ls1\\ilvl0\\cf0 {\\listtext\t\\uc0\\u8226 \t}Item\\\n\n" +
        "\\itap1\\trowd \\taflags0 \\trgaph108\\trleft-108 \n" +
        "\\clvertalc \\clshdrawnil \\clwWidth693\\clftsWidth3 \\clbrdrt\\brdrs\\brdrw20\\brdrcf2 \\clbrdrl\\brdrs\\brdrw20\\brdrcf2 \\clbrdrb\\brdrs\\brdrw20\\brdrcf2 \\clbrdrr\\brdrs\\brdrw20\\brdrcf2 \\clpadt20 \\clpadl20 \\clpadb20 \\clpadr20 \\gaph\\cellx4320\n" +
        "\\clvertalc \\clshdrawnil \\clwWidth506\\clftsWidth3 \\clbrdrt\\brdrs\\brdrw20\\brdrcf2 \\clbrdrl\\brdrs\\brdrw20\\brdrcf2 \\clbrdrb\\brdrs\\brdrw20\\brdrcf2 \\clbrdrr\\brdrs\\brdrw20\\brdrcf2 \\clpadt20 \\clpadl20 \\clpadb20 \\clpadr20 \\gaph\\cellx8640\n" +
        "\\pard\\intbl\\itap1\\pardeftab720\\partightenfactor0\n\\cf0 Region\\cell \n\\pard\\intbl\\itap1\\pardeftab720\\partightenfactor0\n\\cf0 Sales\\cell \\row\n\n" +
        "\\itap1\\trowd \\taflags0 \\trgaph108\\trleft-108 \n" +
        "\\clvertalc \\clshdrawnil \\clwWidth693\\clftsWidth3 \\gaph\\cellx4320\n" +
        "\\clvertalc \\clshdrawnil \\clwWidth506\\clftsWidth3 \\gaph\\cellx8640\n" +
        "\\pard\\intbl\\itap1\\pardeftab720\\partightenfactor0\n\\cf0 North\\cell \n\\pard\\intbl\\itap1\\pardeftab720\\partightenfactor0\n\\cf0 100\\cell \\lastrow\\row\n}";

    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    [InlineData("\r")]
    public void BackslashLineEndingsEndParagraphs(string lineEnding) {
        RtfDocument document = RtfDocument.Parse("{\\rtf1\\ansi First\\" + lineEnding + "Second\\par}");

        Assert.Equal(new[] { "First", "Second" }, document.Blocks.OfType<RtfParagraph>().Select(paragraph => paragraph.ToPlainText()));
    }

    [Fact]
    public void BackslashLineEndingKeepsCellParagraphsApartAndLeavesPrecedingTextOutOfTheTable() {
        RtfDocument document = RtfDocument.Parse(Report);

        Assert.Equal(new[] { "Heading", "Body text", "Item" },
            document.Blocks.OfType<RtfParagraph>().Select(paragraph => paragraph.ToPlainText()));
        RtfTable table = Assert.Single(document.Blocks.OfType<RtfTable>());
        Assert.Equal(new[] { "Region", "Sales" }, table.Rows[0].Cells.Select(CellText));
        Assert.Equal(new[] { "North", "100" }, table.Rows[1].Cells.Select(CellText));

        static string CellText(RtfTableCell cell) => string.Concat(cell.Blocks.OfType<RtfParagraph>().Select(paragraph => paragraph.ToPlainText()));
    }

    [Fact]
    public void UngroupedFontTableEntriesAreRead() {
        RtfDocument document = RtfDocument.Parse(Report);

        Assert.Equal(new[] { "Times-Bold", "Times-Roman" }, document.Fonts.Select(font => font.Name));
        Assert.Equal(new[] { 0, 1 }, document.Fonts.Select(font => font.Id));
        Assert.All(document.Fonts, font => Assert.Equal(RtfFontFamily.Roman, font.Family));
    }

    [Fact]
    public void FontTableReadsUngroupedAndGroupedEntriesTogether() {
        RtfDocument document = RtfDocument.Parse("{\\rtf1\\ansi{\\fonttbl\\f0\\fswiss Helvetica;{\\f1\\froman Georgia;}}\\f1 Text\\par}");

        Assert.Equal(new[] { "Helvetica", "Georgia" }, document.Fonts.Select(font => font.Name));
    }

    [Fact]
    public void ExpandedColorTableIsNotALoss() {
        RtfReadResult result = RtfDocument.ParseResult(Report);

        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF101");
    }

    [Fact]
    public void PdfKeepsParagraphsOutsideTheTableAndAddsNoTableFills() {
        var options = new RtfToPdfOptions();
        using var pdf = PdfPigDocument.Open(RtfDocument.Parse(Report).ConvertToPdfWithEvidence(options).ToBytes());
        var page = pdf.GetPage(1);
        double Baseline(string word) => page.GetWords().Single(candidate => candidate.Text == word).BoundingBox.Bottom;

        Assert.True(Baseline("Heading") > Baseline("Body"));
        Assert.True(Baseline("Body") > Baseline("Item"));
        Assert.True(Baseline("Item") > Baseline("Region"));
        Assert.Equal(Baseline("Region"), Baseline("Sales"), 1);
        Assert.True(Baseline("Region") > Baseline("North"));
        Assert.DoesNotContain(page.Paths, path => path.IsFilled && path.GetBoundingRectangle() is { Height: > 2D, Width: > 2D });
        Assert.DoesNotContain(options.Report.Warnings, warning => warning.Code == "FontFamilySlotExhausted");
    }

    [Fact]
    public void ReadingThePdfKeepsAHangingBulletWithItsText() {
        byte[] bytes = RtfDocument.Parse(Report).ConvertToPdfWithEvidence(new RtfToPdfOptions()).ToBytes();

        PdfCore.PdfLogicalListItem item = Assert.Single(PdfCore.PdfDocument.Load(bytes).Read(new PdfCore.PdfLoadOptions()).Pages[0].ListItems);

        Assert.Equal("\u2022", item.Marker);
        Assert.Equal("Item", item.Text);
    }

    [Fact]
    public void ColumnNarrowerThanItsPaddingStillConverts() {
        byte[] bytes = RtfDocument.Parse("{\\rtf1\\ansi\\trowd\\trgaph108\\cellx14\\cellx3000\\pard\\intbl\\cell Wide\\cell\\row}")
            .ConvertToPdfWithEvidence(new RtfToPdfOptions()).ToBytes();

        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Wide", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void BorderWidthIsReadAsTwips() {
        RtfTable table = Assert.Single(RtfDocument.Parse(Report).Blocks.OfType<RtfTable>());

        PdfCore.PdfCellBorder border = RtfPdfMapping.ToPdfTableStyle(RtfDocument.Parse(Report), table, new RtfToPdfOptions()).CellBorders![(0, 0)];

        Assert.Equal(1D, border.TopBorder!.Width.ToPoints(), 6);
    }

    [Fact]
    public void ListMarkerSharesTheTextLineBoxSoItemsDoNotCollide() {
        using var pdf = PdfPigDocument.Open(RtfDocument.Parse(
            "{\\rtf1\\ansi{\\fonttbl{\\f0\\froman Times New Roman;}}" +
            "{\\*\\listtable{\\list\\listtemplateid1{\\listlevel\\levelnfc23\\leveljc0\\levelfollow0\\levelstartat1{\\leveltext\\'01\\u8226 ?;}{\\levelnumbers;}\\fi-360\\li720}\\listid1}}" +
            "{\\*\\listoverridetable{\\listoverride\\listid1\\listoverridecount0\\ls1}}" +
            "\\pard\\li720\\fi-360\\ls1\\f0\\fs24 Alpha\\par\\pard\\li720\\fi-360\\ls1\\f0\\fs24 Beta\\par}").ConvertToPdfWithEvidence(new RtfToPdfOptions()).ToBytes());
        var words = pdf.GetPage(1).GetWords().ToArray();

        double advance = words.Single(word => word.Text.EndsWith("Alpha", StringComparison.Ordinal)).BoundingBox.Bottom -
            words.Single(word => word.Text.EndsWith("Beta", StringComparison.Ordinal)).BoundingBox.Bottom;
        Assert.True(advance > 12.5D, $"The line advance was {advance}.");
    }

    [Fact]
    public void TabStopsAreMeasuredFromTheMarginInsideALeftIndent() {
        using var pdf = PdfPigDocument.Open(RtfDocument.Parse(
            "{\\rtf1\\ansi\\margl1440\\pard\\li720\\tx2880 A\\tab B\\par}").ConvertToPdfWithEvidence(new RtfToPdfOptions()).ToBytes());
        var letters = pdf.GetPage(1).Letters;

        Assert.Equal(216D, letters.Single(letter => letter.Value == "B").StartBaseLine.X, 0);
    }

    [Fact]
    public void HangingMarkerTabReachesTheLeftIndent() {
        using var pdf = PdfPigDocument.Open(RtfDocument.Parse(
            "{\\rtf1\\ansi\\margl1440\\pard\\tx220\\tx720\\li720\\fi-720 A\\tab B\\par}").ConvertToPdfWithEvidence(new RtfToPdfOptions()).ToBytes());
        var letters = pdf.GetPage(1).Letters;

        Assert.Equal(108D, letters.Single(letter => letter.Value == "B").StartBaseLine.X, 0);
    }
}
