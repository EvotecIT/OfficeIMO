using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfSmallCapsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InheritedSmallCapsRestoreFullSizeAfterLowercaseText(bool hyperlink) {
        using WordDocument document = CreateDocument();
        WordParagraph paragraph = document.AddParagraph();
        if (hyperlink) paragraph.AddHyperLink("aB", new Uri("https://example.com/caps")).Paragraph.CapsStyle = WordCapsStyle.SmallCaps;
        else {
            paragraph.Text = "aB";
            paragraph.SetSmallCaps();
        }
        using var pdf = OpenPdf(document);
        Assert.Equal("AB", pdf.GetPage(1).Text);
        Assert.Equal(new[] { 9.5D, 12D }, pdf.GetPage(1).Letters.Select(letter => letter.PointSize));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OriginalSmallCapsCapitalsRetainConfiguredTableFontSize(bool hyperlink) {
        using WordDocument document = CreateDocument();
        WordTable table = document.AddTable(1, 1);
        table._tableProperties!.TableStyle?.Remove();
        WordParagraph paragraph = table.Rows[0].Cells[0].Paragraphs[0];
        if (hyperlink) paragraph.AddHyperLink("AB", new Uri("https://example.com/caps")).Paragraph.CapsStyle = WordCapsStyle.SmallCaps;
        else {
            paragraph.Text = "AB";
            paragraph.SetSmallCaps();
        }
        WordPdfExportOptions options = Options();
        options.PdfOptions = new PdfCreateOptions {
            DefaultTableStyle = new PdfTableStyle { FontSize = 18D, HeaderRowCount = 0 }
        };
        using var pdf = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal("AB", pdf.GetPage(1).Text);
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.Equal(18D, letter.PointSize, 3));
    }

    [Theory]
    [InlineData(2D, 1.5D)]
    [InlineData(3D, 2.5D)]
    [InlineData(4D, 3D)]
    [InlineData(8.5D, 7D)]
    [InlineData(9D, 7D)]
    [InlineData(9.5D, 7.5D)]
    [InlineData(11D, 9D)]
    [InlineData(12D, 9.5D)]
    [InlineData(12.5D, 10D)]
    [InlineData(22D, 17.5D)]
    [InlineData(30D, 24D)]
    public void SmallCapsUsesNativeHalfPointSizesAndKeepsOriginalCapitals(double fontSize, double smallSize) {
        using WordDocument document = CreateDocument();
        WordParagraph run = document.AddParagraph("Aa 1!");
        run.FontSizePoints = fontSize;
        run.SetSmallCaps();
        using var pdf = OpenPdf(document);
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value != " ").ToArray();
        Assert.Equal("AA1!", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(fontSize, letters[0].PointSize, 3);
        Assert.Equal(smallSize, letters[1].PointSize, 3);
        Assert.Equal(fontSize, letters[2].PointSize, 3);
        Assert.Equal(fontSize, letters[3].PointSize, 3);
        Assert.Equal(smallSize, Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == " ").PointSize, 3);
        Assert.All(letters, letter => Assert.Equal(letters[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
    }

    [Theory]
    [InlineData("body", "direct")]
    [InlineData("cell", "direct")]
    [InlineData("header", "direct")]
    [InlineData("footer", "direct")]
    [InlineData("body", "paragraph")]
    [InlineData("cell", "paragraph")]
    [InlineData("header", "paragraph")]
    [InlineData("footer", "paragraph")]
    [InlineData("body", "character")]
    [InlineData("cell", "character")]
    [InlineData("header", "character")]
    [InlineData("footer", "character")]
    [InlineData("body", "document")]
    [InlineData("cell", "document")]
    [InlineData("header", "document")]
    [InlineData("footer", "document")]
    [InlineData("cell", "table")]
    [InlineData("cell", "conditional")]
    [InlineData("body-textbox", "direct")]
    [InlineData("header-textbox", "character")]
    public void SmallCapsInheritanceAndExplicitResetReachEveryStory(string route, string source) {
        using WordDocument document = CreateDocument();
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        WordTable? table = route == "cell" ? document.AddTable(1, 1) : null;
        WordParagraph paragraph = route switch {
            "cell" => table!.Rows[0].Cells[0].Paragraphs[0],
            "header" => document.HeaderDefaultOrCreate.AddParagraph(),
            "footer" => document.FooterDefaultOrCreate.AddParagraph(),
            "body-textbox" => document.AddParagraph().AddTextBox(string.Empty, WordImageTextWrapping.Square).Paragraphs.Single(),
            "header-textbox" => document.HeaderDefaultOrCreate.AddParagraph().AddTextBox(string.Empty, WordImageTextWrapping.Square).Paragraphs.Single(),
            _ => document.AddParagraph()
        };
        WordParagraph run = paragraph.AddText("Aa");
        WordParagraph reset = paragraph.AddText("Bb");
        if (source == "document") {
            styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!.Append(new W.SmallCaps());
        } else if (source is "paragraph" or "character") {
            styles.Append(new W.Style(new W.StyleRunProperties(new W.SmallCaps())) {
                Type = source == "character" ? W.StyleValues.Character : W.StyleValues.Paragraph,
                StyleId = "InheritedSmallCaps"
            });
            if (source == "paragraph") paragraph.SetStyleId("InheritedSmallCaps");
            else foreach (WordParagraph item in new[] { run, reset })
                (item._run.RunProperties ??= new W.RunProperties()).RunStyle = new W.RunStyle { Val = "InheritedSmallCaps" };
        } else if (source is "table" or "conditional") {
            var tableStyle = new W.Style { Type = W.StyleValues.Table, StyleId = "SmallCapsTable" };
            if (source == "table") tableStyle.Append(new W.StyleRunProperties(new W.SmallCaps()));
            else {
                tableStyle.Append(new W.TableStyleProperties(new W.RunPropertiesBaseStyle(new W.SmallCaps())) {
                    Type = W.TableStyleOverrideValues.FirstRow
                });
                table!.ConditionalFormattingFirstRow = true;
            }
            styles.Append(tableStyle);
            table!.SetStyleId("SmallCapsTable");
        } else run.SetSmallCaps();
        (reset._run.RunProperties ??= new W.RunProperties()).SmallCaps = new W.SmallCaps { Val = false };
        if (route is "header" or "footer" or "header-textbox") document.AddParagraph();
        string before = document._wordprocessingDocument.MainDocumentPart.Document.OuterXml;
        using var pdf = OpenPdf(document);
        var letters = pdf.GetPage(1).Letters.ToArray();
        Assert.Equal("AABb", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(new[] { 12D, 9.5D, 12D, 12D }, letters.Select(letter => letter.PointSize).ToArray());
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Fact]
    public void SmallCapsMeasuresReducedGlyphsBeforeWrappingMixedCaseWords() {
        using WordDocument document = CreateDocument();
        document.AddParagraph("MMMMmmmm Z").SetSmallCaps();
        using var pdf = OpenPdf(document, new PdfPageSize(115, 300));
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value != " ").ToArray();
        Assert.Equal("MMMMMMMMZ", string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters.Take(8), letter => Assert.Equal(letters[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        Assert.True(letters[8].StartBaseLine.Y < letters[0].StartBaseLine.Y);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SmallCapsOffDoesNotDisableInheritedAllCaps(bool table) {
        using WordDocument document = CreateDocument();
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new W.Style(new W.StyleRunProperties(new W.Caps())) {
            Type = W.StyleValues.Paragraph, StyleId = "AllCapsParagraph"
        });
        WordParagraph paragraph = table ? document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] : document.AddParagraph();
        paragraph.SetStyleId("AllCapsParagraph");
        WordParagraph run = paragraph.AddText("Aa");
        (run._run.RunProperties ??= new W.RunProperties()).SmallCaps = new W.SmallCaps { Val = false };
        using var pdf = OpenPdf(document);
        Assert.Equal("AA", pdf.GetPage(1).Text);
        Assert.All(pdf.GetPage(1).Letters, letter => Assert.Equal(12D, letter.PointSize, 3));
    }

    [Fact]
    public void SmallCapsPreservesHyperlinkDestinationAndTooltip() {
        using WordDocument document = CreateDocument();
        const string uri = "https://example.com/small-caps";
        WordParagraph paragraph = document.AddParagraph();
        paragraph.AddHyperLink("AaBb", new Uri(uri), tooltip: "Small caps reference").Paragraph.CapsStyle = WordCapsStyle.SmallCaps;
        byte[] bytes = document.ToPdf(Options()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal("AABB", pdf.GetPage(1).Text);
        Assert.Equal(new[] { 12D, 9.5D, 12D, 9.5D }, pdf.GetPage(1).Letters.Select(letter => letter.PointSize));
        var links = PdfDocumentReadResult.Load(bytes).GetLinksByUri(uri).ToArray();
        Assert.NotEmpty(links);
        Assert.All(links, link => Assert.Equal("Small caps reference", link.Contents));
    }

    private static WordDocument CreateDocument() {
        WordDocument document = WordDocument.Create();
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.DocDefaults = new W.DocDefaults(new W.RunPropertiesDefault(new W.RunPropertiesBaseStyle(
            new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "24" })));
        W.Style normal = styles.Elements<W.Style>().Single(style => style.StyleId?.Value == "Normal");
        normal.StyleRunProperties = new W.StyleRunProperties();
        normal.StyleParagraphProperties = new W.StyleParagraphProperties(new W.SpacingBetweenLines {
            Before = "0", After = "0", Line = "240", LineRule = W.LineSpacingRuleValues.Auto
        });
        return document;
    }

    private static WordPdfExportOptions Options(PdfPageSize? pageSize = null) => new() {
        IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
        PageSize = pageSize, Margins = PdfPageMargins.Uniform(20)
    };

    private static PdfPigDocument OpenPdf(WordDocument document, PdfPageSize? pageSize = null) =>
        PdfPigDocument.Open(document.ToPdf(Options(pageSize)).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
}
