using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordListMarkerSemanticsTests {
    private static string IssueDocumentPath => Path.Combine(AppContext.BaseDirectory, "Documents", "Issue2510-SymbolBullets.docx");

    [Fact]
    public void WordAuthoredSymbolBulletsKeepTheirKindAndPortableMarkers() {
        using WordDocument document = WordDocument.Load(IssueDocumentPath);
        WordParagraph[] items = document.Paragraphs.Where(paragraph => paragraph.IsListItem).ToArray();

        Assert.Equal(2, items.Length);
        Assert.All(items, paragraph => Assert.False(WordDocumentTraversal.GetListInfo(paragraph)!.Value.Ordered));
        Assert.All(WordDocumentTraversal.BuildListMarkers(document).Values, marker => Assert.Equal("•", marker.Marker));
        Assert.Equal("- Test\n- Test", document.ToMarkdown().Trim());

        WordDocumentVisualSnapshot snapshot = document.CreateVisualSnapshot();
        OfficeDrawingText firstBody = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().First(text => text.Text == "Test");
        OfficeDrawingText markerText = snapshot.Drawing.Elements.OfType<OfficeDrawingText>()
            .Single(text => text.Y == firstBody.Y && text.X < firstBody.X);
        Assert.Equal("•", markerText.Text);
        Assert.NotEqual("Symbol", markerText.Font.FamilyName);
        string svg = Encoding.UTF8.GetString(document.ExportImage(OfficeImageExportFormat.Svg).Bytes);
        Assert.Contains("•", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("\uf0b7", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void WordAuthoredNestedWingdingsMarkerIsProjectedWithoutPrivateUseCodepoint() {
        using WordDocument document = WordDocument.Load(IssueDocumentPath);
        WordParagraph nestedItem = document.Paragraphs.First(paragraph => paragraph.IsListItem);
        nestedItem._paragraph.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val = 2;

        WordDocumentTraversal.ListInfo info = WordDocumentTraversal.GetListInfo(nestedItem)!.Value;
        Assert.False(info.Ordered);
        Assert.Equal("Wingdings", info.MarkerFontFamily);
        Assert.Equal("▪", WordDocumentTraversal.BuildListMarkers(document)[nestedItem].Marker);
        string pdfText = PdfReadDocument.Open(document.ToPdfBytes()).ExtractText();
        Assert.Contains("▪", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("\uf0a7", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void WordAuthoredBulletsExportAsUnorderedHtmlAndReadablePdf() {
        using WordDocument document = WordDocument.Load(IssueDocumentPath);
        string html = document.ToHtml();
        Assert.Contains("<ul", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<ol", html, StringComparison.OrdinalIgnoreCase);

        using var output = new MemoryStream();
        document.SaveAsPdfResult(output);
        byte[] pdf = output.ToArray();
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Equal(2, text.Split("Test", StringSplitOptions.None).Length - 1);
        Assert.DoesNotContain("\uf0b7", text, StringComparison.Ordinal);
        Assert.Contains("•", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfIncludesMarkersForListParagraphsInTableCellTextBoxAndHeader() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.BulletSolidRound));
        WordTable table = document.AddTable(1, 1);
        WordParagraph cell = table.Rows[0].Cells[0].Paragraphs[0];
        cell.Text = "Cell marker";
        AttachToList(cell, list.NumberId);
        WordTextBox textBox = document.AddTextBox("Box marker");
        AttachToList(textBox.Paragraphs[0], list.NumberId);
        WordParagraph header = document.Header!.Default!.AddParagraph("Header marker");
        AttachToList(header, list.NumberId);

        Dictionary<WordParagraph, (int Level, string Marker)> markers = WordDocumentTraversal.BuildListMarkers(document);
        Assert.Equal("•", markers[cell].Marker);
        Assert.Equal("•", markers[textBox.Paragraphs[0]].Marker);
        Assert.Equal("•", markers[header].Marker);
        string pdfText = PdfReadDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false })).ExtractText();
        Assert.Contains("• Cell marker", pdfText, StringComparison.Ordinal);
        Assert.Contains("• Box marker", pdfText, StringComparison.Ordinal);
        Assert.Contains("• Header marker", pdfText, StringComparison.Ordinal);
    }

    private static void AttachToList(WordParagraph paragraph, int numberId) {
        paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph._paragraph.ParagraphProperties.NumberingProperties = new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = numberId });
    }

    [Fact]
    public void FullLevelOverrideWinsOverAbstractFormatAndMarkerFormatting() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        WordParagraph item = list.AddItem("Overridden bullet");
        Numbering numbering = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
        NumberingInstance instance = numbering.Elements<NumberingInstance>()
            .Single(candidate => candidate.NumberID?.Value == list.NumberId);
        instance.Elements<LevelOverride>().Single(level => level.LevelIndex?.Value == 0).Remove();
        instance.Append(new LevelOverride(new Level(
            new StartNumberingValue { Val = 7 },
            new NumberingFormat { Val = NumberFormatValues.Bullet },
            new LevelText { Val = "◆" },
            new LevelJustification { Val = LevelJustificationValues.Right },
            new LevelSuffix { Val = LevelSuffixValues.Space },
            new PreviousParagraphProperties(new Indentation { Left = "900", Hanging = "300" }),
            new NumberingSymbolRunProperties(new RunFonts { Ascii = "Arial" }, new Bold(), new Color { Val = "FF0000" })) {
            LevelIndex = 0
        }) { LevelIndex = 0 });

        WordDocumentTraversal.ListInfo info = WordDocumentTraversal.GetListInfo(item)!.Value;
        Assert.False(info.Ordered);
        Assert.Equal(WordNumberFormat.Bullet, info.NumberFormat);
        Assert.Equal(7, info.Start);
        Assert.Equal("◆", info.LevelText);
        Assert.Equal("Arial", info.MarkerFontFamily);
        Assert.Equal("FF0000", info.MarkerColorHex);
        Assert.Equal(900, info.LeftIndentTwips);
        Assert.Equal(300, info.HangingIndentTwips);
        Assert.Equal(WordListLevelAlignment.Right, info.LevelJustification);
        Assert.Equal(WordListLevelSuffix.Space, info.LevelSuffix);
        Assert.Equal("◆", WordDocumentTraversal.BuildListMarkers(document)[item].Marker);
    }

    [Fact]
    public void NumberingInheritedThroughStylesCanBeCancelledByDirectZeroId() {
        using WordDocument document = WordDocument.Create();
        WordList bullets = document.AddCustomBulletList('◆', "Arial", "000000");
        bullets.AddItem("Seed");
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(new StyleParagraphProperties(new NumberingProperties(
            new NumberingLevelReference { Val = 0 },
            new NumberingId { Val = bullets.NumberId }))) {
            Type = StyleValues.Paragraph,
            StyleId = "Issue2510BaseList"
        });
        styles.Append(new Style(new BasedOn { Val = "Issue2510BaseList" }) {
            Type = StyleValues.Paragraph,
            StyleId = "Issue2510InheritedList"
        });
        WordParagraph inherited = document.AddParagraph("Inherited bullet");
        inherited._paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Issue2510InheritedList" });
        WordParagraph cancelled = document.AddParagraph("Plain paragraph");
        cancelled._paragraph.ParagraphProperties = new ParagraphProperties(
            new ParagraphStyleId { Val = "Issue2510InheritedList" },
            new NumberingProperties(new NumberingId { Val = 0 }));

        Assert.True(inherited.IsListItem);
        Assert.False(WordDocumentTraversal.GetListInfo(inherited)!.Value.Ordered);
        Assert.Equal("◆", WordDocumentTraversal.BuildListMarkers(document)[inherited].Marker);
        Assert.False(cancelled.IsListItem);
        Assert.Null(WordDocumentTraversal.GetListInfo(cancelled));
        Assert.DoesNotContain(cancelled, WordDocumentTraversal.BuildListMarkers(document).Keys);
    }

    [Fact]
    public void MixedAndUnmarkedLevelsFollowEachEffectiveNumberFormat() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.BulletSquareSymbol));
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));
        WordParagraph numbered = list.AddItem("One", 0);
        WordParagraph bullet = list.AddItem("Square", 1);
        WordParagraph unmarked = list.AddItem("No marker", 2);

        Assert.True(WordDocumentTraversal.GetListInfo(numbered)!.Value.Ordered);
        Assert.False(WordDocumentTraversal.GetListInfo(bullet)!.Value.Ordered);
        Assert.False(WordDocumentTraversal.GetListInfo(unmarked)!.Value.MarkerVisible);
        Assert.Equal("1.", WordDocumentTraversal.BuildListMarkers(document)[numbered].Marker);
        Assert.Equal("■", WordDocumentTraversal.BuildListMarkers(document)[bullet].Marker);
        Assert.Equal(string.Empty, WordDocumentTraversal.BuildListMarkers(document)[unmarked].Marker);
        string pdfText = PdfReadDocument.Open(document.ToPdfBytes()).ExtractText();
        Assert.Contains("No marker", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("• No marker", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void PictureBulletRetainsItsIdentityWithDeterministicTextFallback() {
        using WordDocument document = WordDocument.Create();
        using var image = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg"));
        WordList list = document.AddPictureBulletList(image, "Kulek.jpg");
        WordParagraph item = list.AddItem("Picture bullet");

        WordDocumentTraversal.ListInfo info = WordDocumentTraversal.GetListInfo(item)!.Value;
        Assert.False(info.Ordered);
        Assert.True(info.PictureBulletId > 0);
        Assert.Equal("•", WordDocumentTraversal.BuildListMarkers(document)[item].Marker);
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        Assert.Contains(result.Warnings, warning => warning.Code == "NativePictureBulletTextFallback");
        Assert.Contains("•", PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }
}
