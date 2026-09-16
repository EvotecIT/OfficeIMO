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
        string html = document.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });
        Assert.DoesNotContain("\uf0a7", html, StringComparison.Ordinal);
        Assert.Contains("list-style-type:'▪'", html, StringComparison.Ordinal);
    }

    [Fact]
    public void WordAuthoredBulletsExportAsUnorderedHtmlAndReadablePdf() {
        using WordDocument document = WordDocument.Load(IssueDocumentPath);
        string html = document.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });
        Assert.Contains("<ul", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<ol", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("\uf0b7", html, StringComparison.Ordinal);
        Assert.Contains("list-style-type:disc", html, StringComparison.OrdinalIgnoreCase);

        using var output = new MemoryStream();
        document.SaveAsPdfResult(output);
        byte[] pdf = output.ToArray();
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Equal(2, text.Split(new[] { "Test" }, StringSplitOptions.None).Length - 1);
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

    [Fact]
    public void PdfIncludesMarkersInsideHeaderTextBoxes() {
        using WordDocument document = WordDocument.Create();
        document.AddHeadersAndFooters();
        WordList list = document.AddCustomBulletList('◆', "Arial", "000000");
        WordParagraph host = document.Header!.Default!.AddParagraph("Before ");
        WordTextBox box = host.AddTextBox("Header box marker", WordImageTextWrapping.Square);
        host.AddText(" After");
        AttachToList(box.Paragraphs[0], list.NumberId);
        document.AddParagraph("Body");

        string pdfText = PdfReadDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false })).ExtractText();
        Assert.Contains("Before ◆ Header box marker", pdfText, StringComparison.Ordinal);
        Assert.Contains("After", pdfText, StringComparison.Ordinal);
        Assert.True(pdfText.IndexOf("Header box marker", StringComparison.Ordinal) < pdfText.IndexOf("After", StringComparison.Ordinal));
    }

    [Fact]
    public void PdfTableMarkersUseLevelFormatting() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomBulletList('*', "Arial", "FF0000");
        WordListLevel level = list.Numbering.Levels[0];
        level.OpenXmlElement.GetFirstChild<NumberingSymbolRunProperties>()!.Append(new Bold(), new FontSize { Val = "36" });
        level.IndentationLeft = 2160;
        level.IndentationHanging = 360;
        WordTable table = document.AddTable(1, 1);
        table.Width = 7000;
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.Width = 7000;
        cell.Paragraphs[0].Text = "PlainCell";
        AttachToList(cell.AddParagraph("CellStyle"), list.NumberId);

        var spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans();
        PdfTextSpan marker = Assert.Single(spans, span => span.Text.Contains("*", StringComparison.Ordinal));
        PdfTextSpan plain = Assert.Single(spans, span => span.Text == "PlainCell");
        Assert.True(marker.X > plain.X + 45D);
        Assert.InRange(marker.FontSize, 17.5D, 18.5D);
        Assert.True(marker.IsBold);
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), marker.Color);
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
    public void FullLevelOverrideDoesNotInheritOmittedAbstractProperties() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomBulletList('\uf0b7', "Symbol", "FF0000");
        WordParagraph item = list.AddItem("Override replacement");
        Level abstractLevel = list.Numbering.Levels[0].OpenXmlElement;
        abstractLevel.GetFirstChild<PreviousParagraphProperties>()!.GetFirstChild<Indentation>()!.Left = "1800";
        Numbering numbering = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
        NumberingInstance instance = numbering.Elements<NumberingInstance>()
            .Single(candidate => candidate.NumberID?.Value == list.NumberId);
        foreach (LevelOverride oldOverride in instance.Elements<LevelOverride>().ToArray()) oldOverride.Remove();
        instance.Append(new LevelOverride(new Level(
            new NumberingFormat { Val = NumberFormatValues.Bullet },
            new LevelText { Val = "◆" }) { LevelIndex = 0 }) { LevelIndex = 0 });

        WordDocumentTraversal.ListInfo info = WordDocumentTraversal.GetListInfo(item)!.Value;
        Assert.Equal("◆", WordDocumentTraversal.BuildListMarkers(document)[item].Marker);
        Assert.Null(info.MarkerFontFamily);
        Assert.Null(info.MarkerColorHex);
        Assert.Null(info.LeftIndentTwips);
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
    public void SettingInheritedListLevelCreatesAnEffectiveDirectOverride() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.BulletSolidRound));
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(new StyleParagraphProperties(new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = list.NumberId }))) {
            Type = StyleValues.Paragraph, StyleId = "Issue2510SetterStyle"
        });
        WordParagraph item = document.AddParagraph("Changed level");
        item._paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Issue2510SetterStyle" });

        item.ListItemLevel = 1;
        Assert.Equal(1, item.ListItemLevel);
        Assert.Equal(list.NumberId, item._paragraph.ParagraphProperties.NumberingProperties!.NumberingId!.Val!.Value);
        Assert.False(WordDocumentTraversal.GetListInfo(item)!.Value.Ordered);
        item.ListItemLevel = null;
        Assert.Equal(0, item.ListItemLevel);
    }

    [Fact]
    public void HtmlPreservesDefaultStyleNumberingOnPlainParagraphs() {
        using WordDocument document = WordDocument.Create();
        WordList bullets = document.AddCustomBulletList('◆', "Arial", "000000");
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style defaultStyle = styles.Elements<Style>()
            .Single(style => style.Type?.Value == StyleValues.Paragraph && style.Default?.Value == true);
        defaultStyle.StyleParagraphProperties ??= new StyleParagraphProperties();
        defaultStyle.StyleParagraphProperties.NumberingProperties = new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = bullets.NumberId });

        WordParagraph item = document.AddParagraph("Default style bullet");
        item._paragraph.ParagraphProperties?.Remove();

        Assert.True(item.IsListItem);
        Assert.Null(item._paragraph.ParagraphProperties);
        string html = document.ToHtml();
        Assert.Contains("<ul", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<li", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Default style bullet", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<p>Default style bullet</p>", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void StyleLinkedNumberingLevelOverridesStyleNumPrLevel() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.BulletSquareSymbol));
        list.Numbering.Levels[2].OpenXmlElement.Append(new ParagraphStyleIdInLevel { Val = "Issue2510LinkedLevel" });
        Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new Style(new StyleParagraphProperties(new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = list.NumberId }))) {
            Type = StyleValues.Paragraph, StyleId = "Issue2510LinkedLevel"
        });
        WordParagraph item = document.AddParagraph("Linked square");
        item._paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Issue2510LinkedLevel" });

        WordDocumentTraversal.ListInfo info = WordDocumentTraversal.GetListInfo(item)!.Value;
        Assert.Equal(2, info.Level);
        Assert.False(info.Ordered);
        Assert.Equal("■", WordDocumentTraversal.BuildListMarkers(document)[item].Marker);
        Assert.Contains("- Linked square", document.ToMarkdown(), StringComparison.Ordinal);
        Assert.Contains("<ul", document.ToHtml(), StringComparison.OrdinalIgnoreCase);

        Numbering numbering = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
        NumberingInstance instance = numbering.Elements<NumberingInstance>()
            .Single(candidate => candidate.NumberID?.Value == list.NumberId);
        instance.Elements<LevelOverride>().Single(level => level.LevelIndex?.Value == 2).Remove();
        instance.Append(new LevelOverride(new Level(
            new NumberingFormat { Val = NumberFormatValues.Decimal },
            new LevelText { Val = "%3." }) { LevelIndex = 2 }) { LevelIndex = 2 });
        Assert.Equal(0, WordDocumentTraversal.GetListInfo(item)!.Value.Level);
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
        string markdown = document.ToMarkdown();
        Assert.Contains("No marker", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("- No marker", markdown, StringComparison.Ordinal);
        string html = document.ToHtml();
        Assert.Contains("list-style-type:none", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkerlessMarkdownContentStaysWithItsParentListItem() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.BulletSolidRound));
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.BulletSolidRound));
        list.AddItem("Parent", 0);
        list.AddItem("Unmarked child", 1);
        list.AddItem("Marked grandchild", 2);

        string markdown = document.ToMarkdown();
        Assert.Contains("Parent", markdown, StringComparison.Ordinal);
        Assert.Contains("Unmarked child", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("- Unmarked child", markdown, StringComparison.Ordinal);
        Assert.Contains("  Unmarked child", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain(markdown.Split('\n'), line => line.Trim() == "-");
        Assert.Contains("  - Marked grandchild", markdown, StringComparison.Ordinal);
        Assert.True(markdown.IndexOf("Unmarked child", StringComparison.Ordinal) < markdown.IndexOf("Marked grandchild", StringComparison.Ordinal));
    }

    [Fact]
    public void MarkerlessPdfItemsKeepNumberingLevelIndentation() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.None));
        list.Numbering.Levels[0].IndentationLeft = 2160;
        list.Numbering.Levels[0].IndentationHanging = 360;
        document.AddParagraph("PlainBody");
        list.AddItem("UnmarkedBody");
        WordTable table = document.AddTable(1, 1);
        table.Width = 7000;
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.Width = 7000;
        cell.Paragraphs[0].Text = "PlainCell";
        WordParagraph unmarkedCell = cell.AddParagraph("UnmarkedCell");
        AttachToList(unmarkedCell, list.NumberId);

        var spans = PdfReadDocument.Open(document.ToPdfBytes()).Pages[0].GetTextSpans();
        double plainBodyX = Assert.Single(spans, span => span.Text == "PlainBody").X;
        double unmarkedBodyX = Assert.Single(spans, span => span.Text == "UnmarkedBody").X;
        double plainCellX = Assert.Single(spans, span => span.Text == "PlainCell").X;
        double unmarkedCellX = Assert.Single(spans, span => span.Text == "UnmarkedCell").X;
        Assert.True(unmarkedBodyX > plainBodyX + 45D);
        Assert.True(unmarkedCellX > plainCellX + 45D);
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
