using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_NestedTable_PreservesInnerColumnsOnTheSameRow() {
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        WordTableCell cell = outer.Rows[0].Cells[0];
        cell.Paragraphs[0].Text = "BEFOREINNER";
        WordTable inner = cell.AddTable(1, 2, WordTableStyle.TableGrid);
        inner.Rows[0].Cells[0].Paragraphs[0].Text = "INNERLEFT";
        inner.Rows[0].Cells[1].Paragraphs[0].Text = "INNERRIGHT";
        cell.AddParagraph("AFTERINNER");
        string sourceXml = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;

        var result = document.ToPdfDocumentResult();
        using PdfPigDocument pdf = PdfPigDocument.Open(result.ToBytes());
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToList();
        var left = Assert.Single(words, word => word.Text == "INNERLEFT");
        var right = Assert.Single(words, word => word.Text == "INNERRIGHT");
        var before = Assert.Single(words, word => word.Text == "BEFOREINNER");
        var after = Assert.Single(words, word => word.Text == "AFTERINNER");

        Assert.Equal(sourceXml, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Equal(left.BoundingBox.Bottom, right.BoundingBox.Bottom, precision: 2);
        Assert.True(right.BoundingBox.Left > left.BoundingBox.Right,
            "Distinct inner columns must retain their horizontal placement.");
        Assert.True(before.BoundingBox.Bottom > left.BoundingBox.Top,
            "The source paragraph before the inner table must stay above its cells.");
        Assert.True(after.BoundingBox.Top < left.BoundingBox.Bottom,
            "The source paragraph after the inner table must stay below its cells.");
        Assert.DoesNotContain(result.Warnings,
            warning => warning.Code == "NativeNestedTableLayoutApproximated");
    }

    [Fact]
    public void SaveAsPdf_NestedTable_PercentageWidthUsesThePaddedParentCellFrame() {
        const short firstPaddingTwips = 200, secondPaddingTwips = 400;
        // Both authored margins exceed the grid border's interior clearance;
        // zero padding exercises that separate minimum-inset contract instead.
        var first = MeasureNestedTableFrame(parentPaddingTwips: firstPaddingTwips, innerWidthPercentage: 5000);
        var second = MeasureNestedTableFrame(parentPaddingTwips: secondPaddingTwips, innerWidthPercentage: 5000);
        var half = MeasureNestedTableFrame(parentPaddingTwips: secondPaddingTwips, innerWidthPercentage: 2500);

        Assert.Equal((secondPaddingTwips - firstPaddingTwips) / 20D,
            second.BeforeX - first.BeforeX, precision: 2);
        Assert.Equal(second.BeforeX - first.BeforeX, second.InnerX - first.InnerX, precision: 2);
        Assert.True(second.ColumnDistance < first.ColumnDistance,
            "Parent cell padding must reduce the frame used by a percentage-width inner table.");
        Assert.Equal(second.InnerX, half.InnerX, precision: 2);
        Assert.Equal(second.ColumnDistance / 2D, half.ColumnDistance, precision: 2);
    }

    [Fact]
    public void SaveAsPdf_NestedTable_LongRowContinuesEachCellWithoutRepeatingFinishedPeers() {
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(1, 2, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(outer, 6000, 4000, 2000);
        WordTableCell host = outer.Rows[0].Cells[0];
        ConfigureNestedLayoutParagraph(host.Paragraphs[0], "NESTBEFORE");
        ConfigureNestedLayoutParagraph(outer.Rows[0].Cells[1].Paragraphs[0], "OUTERPEER");
        WordTable inner = host.AddTable(1, 2, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(inner, 3600, 2400, 1200);
        string[] tokens = Enumerable.Range(1, 36).Select(index => $"FLOW{index:000}").ToArray();
        for (int index = 0; index < tokens.Length; index++) {
            WordParagraph paragraph = index == 0
                ? inner.Rows[0].Cells[0].Paragraphs[0]
                : inner.Rows[0].Cells[0].AddParagraph();
            ConfigureNestedLayoutParagraph(paragraph, tokens[index]);
        }
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[1].Paragraphs[0], "INNERPEER");
        ConfigureNestedLayoutParagraph(host.AddParagraph(), "NESTAFTER");
        string original = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;

        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false,
            PageSize = new PageSize(400D, 240D),
            Margins = PageMargins.Uniform(30D),
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.True(pdf.NumberOfPages > 1, "The long nested row must exercise a real page continuation.");
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        Assert.Equal(tokens, words.Where(word => word.Text.StartsWith("FLOW", StringComparison.Ordinal))
            .Select(word => word.Text).ToArray());
        foreach (string token in tokens.Concat(new[] { "NESTBEFORE", "NESTAFTER", "INNERPEER", "OUTERPEER" }))
            Assert.Single(words, word => word.Text == token);
        var first = pdf.GetPages().First(page => page.GetWords().Any(word => word.Text == tokens[0]));
        var final = pdf.GetPages().Single(page => page.GetWords().Any(word => word.Text == tokens[tokens.Length - 1]));
        Assert.True(final.Number > first.Number);
        var beforePage = pdf.GetPages().Single(page => page.GetWords().Any(word => word.Text == "NESTBEFORE"));
        Assert.True(beforePage.Number <= first.Number);
        if (beforePage.Number == first.Number)
            Assert.True(Assert.Single(beforePage.GetWords(), word => word.Text == "NESTBEFORE").BoundingBox.Bottom >
                Assert.Single(first.GetWords(), word => word.Text == tokens[0]).BoundingBox.Top);
        var afterPage = pdf.GetPages().Single(page => page.GetWords().Any(word => word.Text == "NESTAFTER"));
        Assert.True(afterPage.Number >= final.Number);
        if (afterPage.Number == final.Number)
            Assert.True(Assert.Single(afterPage.GetWords(), word => word.Text == "NESTAFTER").BoundingBox.Top <
                Assert.Single(final.GetWords(), word => word.Text == tokens[tokens.Length - 1]).BoundingBox.Bottom);
        foreach (var page in pdf.GetPages()) {
            Assert.NotEmpty(page.GetWords());
            foreach (var word in page.GetWords()) {
                Assert.InRange(word.BoundingBox.Left, 0D, page.Width);
                Assert.InRange(word.BoundingBox.Right, 0D, page.Width);
                Assert.InRange(word.BoundingBox.Bottom, 0D, page.Height);
                Assert.InRange(word.BoundingBox.Top, 0D, page.Height);
            }
        }
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeNestedTableLayoutApproximated");
        Assert.Equal(original, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_NestedNoteTableRetainsColumnsOrderAndNativeNotePlacement(bool endnote) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote);
        note.RemoveAllChildren();
        W.Table outer = CreateNoteTable(new[] { new[] { "CELLBEFORE", "OUTERNOTEPEER" } }, repeatHeader: false);
        W.TableCell cell = outer.Elements<W.TableRow>().Single().Elements<W.TableCell>().First();
        cell.Append(CreateNoteTable(new[] { new[] { "NESTNOTELEFT", "NESTNOTERIGHT" } }, repeatHeader: false),
            new W.Paragraph(RichNoteRun("CELLAFTER")));
        note.Append(new W.Paragraph(NoteBodyMarker(endnote), RichNoteRun("NOTEBEFORE")), outer,
            new W.Paragraph(RichNoteRun("NOTEAFTER")));
        Assert.Empty(source.ValidateDocument());
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var main = document._wordprocessingDocument.MainDocumentPart!;
        OpenXmlCompositeElement notes = endnote ? main.EndnotesPart!.Endnotes! : main.FootnotesPart!.Footnotes!;
        string original = main.Document.OuterXml + notes.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml;

        PdfDocumentConversionResult result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var letters = pdf.GetPages().SelectMany(page => page.Letters).ToArray();
        var left = FindRichNoteToken(letters, "NESTNOTELEFT");
        var right = FindRichNoteToken(letters, "NESTNOTERIGHT");
        Assert.Equal(left[0].StartBaseLine.Y, right[0].StartBaseLine.Y, precision: 2);
        Assert.True(right[0].StartBaseLine.X > left[left.Length - 1].EndBaseLine.X);
        foreach (string token in new[] { "NOTEBEFORE", "CELLBEFORE", "CELLAFTER", "OUTERNOTEPEER", "NOTEAFTER" })
            FindRichNoteToken(letters, token);
        var before = FindRichNoteToken(letters, "CELLBEFORE");
        var after = FindRichNoteToken(letters, "CELLAFTER");
        Assert.True(before[0].StartBaseLine.Y > left[0].StartBaseLine.Y);
        Assert.True(after[0].StartBaseLine.Y < left[0].StartBaseLine.Y);
        string text = string.Concat(letters.Select(letter => letter.Value));
        Assert.True(text.IndexOf("NOTEBEFORE", StringComparison.Ordinal) < text.IndexOf("CELLBEFORE", StringComparison.Ordinal));
        Assert.True(text.IndexOf("CELLAFTER", StringComparison.Ordinal) < text.IndexOf("NOTEAFTER", StringComparison.Ordinal));
        Assert.True(FindRichNoteToken(letters, "LINKTARGET")[0].StartBaseLine.Y > left[0].StartBaseLine.Y);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code is "NativeNestedTableLayoutApproximated" or
            "NativeFootnotePlacementApproximated" or "NativeEndnoteContentApproximated" or "NativeEndnoteContinuationApproximated");
        Assert.Equal(original, main.Document.OuterXml + notes.OuterXml + main.StyleDefinitionsPart.Styles.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SaveAsPdf_NestedTableInRowSpanRetainsContentAndReportsTheUnsupportedLayout() {
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(2, 2, WordTableStyle.TableGrid);
        outer.Rows[0].Cells[0].VerticalMerge = WordCellMerge.Restart;
        outer.Rows[1].Cells[0].VerticalMerge = WordCellMerge.Continue;
        WordTable inner = outer.Rows[0].Cells[0].AddTable(1, 2, WordTableStyle.TableGrid);
        inner.Rows[0].Cells[0].Paragraphs[0].Text = "FALLBACKLEFT";
        inner.Rows[0].Cells[1].Paragraphs[0].Text = "FALLBACKRIGHT";
        outer.Rows[0].Cells[1].Paragraphs[0].Text = "UPPERPEER";
        outer.Rows[1].Cells[1].Paragraphs[0].Text = "LOWERPEER";
        string original = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;

        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        foreach (string token in new[] { "FALLBACKLEFT", "FALLBACKRIGHT", "UPPERPEER", "LOWERPEER" })
            Assert.Single(words, word => word.Text == token);
        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code == "NativeNestedTableLayoutApproximated");
        Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.True(result.HasLoss);
        Assert.Equal(original, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_NestedTableAfterWrappedPicturePreservesContentAndReportsLayoutFallback() {
        using WordDocument document = WordDocument.Create();
        WordParagraph anchorParagraph = document.AddParagraph();
        ConfigureNestedLayoutParagraph(anchorParagraph, "PICTUREANCHOR");
        using var pixels = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(4, 2));
        WordImage image = anchorParagraph.AddText(string.Empty).InsertImage(
            pixels, "nested-wrap.png", 80D, 200D, WordImageTextWrapping.Square);
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Column;
        image.HorizontalPositionOffset = 120L * 12700L;
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Paragraph;
        image.VerticalPositionOffset = 0L;
        DW.Anchor pictureAnchor = image._Image.Anchor!;
        pictureAnchor.DistanceFromLeft = pictureAnchor.DistanceFromRight = 0U;
        pictureAnchor.DistanceFromTop = pictureAnchor.DistanceFromBottom = 0U;
        pictureAnchor.BehindDoc = false;
        foreach (DW.EffectExtent extent in pictureAnchor.Descendants<DW.EffectExtent>())
            extent.LeftEdge = extent.TopEdge = extent.RightEdge = extent.BottomEdge = 0L;

        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(outer, 6000, 6000);
        WordTableCell host = outer.Rows[0].Cells[0];
        ConfigureNestedLayoutParagraph(host.Paragraphs[0], "WRAPBEFORE");
        WordTable inner = host.AddTable(1, 2, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(inner, 3600, 1800, 1800);
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[0].Paragraphs[0], "WRAPLEFT");
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[1].Paragraphs[0], "WRAPRIGHT");
        ConfigureNestedLayoutParagraph(host.AddParagraph(), "WRAPAFTER");
        ConfigureNestedLayoutParagraph(document.AddParagraph(), "BODYAFTER");
        Assert.Empty(document.ValidateDocument());
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string original = main.Document.OuterXml + main.StyleDefinitionsPart!.Styles!.OuterXml +
            main.DocumentSettingsPart!.Settings!.OuterXml;

        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false,
            PageSize = new PageSize(400D, 500D),
            Margins = PageMargins.Uniform(30D),
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        PdfConversionWarning warning = Assert.Single(result.Warnings,
            item => item.Code == "NativeNestedTableLayoutApproximated");
        Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var words = pdf.GetPages().SelectMany(page => page.GetWords()).ToArray();
        foreach (string token in new[] { "PICTUREANCHOR", "WRAPBEFORE", "WRAPLEFT", "WRAPRIGHT", "WRAPAFTER", "BODYAFTER" })
            Assert.Single(words, word => word.Text == token);
        var placement = Assert.Single(pdf.GetPages().SelectMany(page => page.GetImages()));
        Assert.Equal(60D, placement.BoundingBox.Width, precision: 2);
        Assert.Equal(150D, placement.BoundingBox.Height, precision: 2);
        Assert.DoesNotContain(result.Warnings, item => item.Code is "NativeAnchoredImageFlowed" or
            "NativeAnchoredImageWrapApproximated");
        Assert.Single(result.Warnings, item => item.Code == "NativeNestedTableLayoutApproximated");
        Assert.True(result.HasLoss);
        Assert.Equal(original, main.Document.OuterXml + main.StyleDefinitionsPart.Styles.OuterXml +
            main.DocumentSettingsPart.Settings.OuterXml);
    }

    private static (double BeforeX, double InnerX, double ColumnDistance) MeasureNestedTableFrame(
        short parentPaddingTwips, int innerWidthPercentage) {
        using WordDocument document = WordDocument.Create();
        WordTable outer = document.AddTable(1, 1, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(outer, 6000, 6000);
        WordTableCell parent = outer.Rows[0].Cells[0];
        parent.MarginLeftWidth = parentPaddingTwips;
        parent.MarginRightWidth = parentPaddingTwips;
        ConfigureNestedLayoutParagraph(parent.Paragraphs[0], "AHEAD");
        WordTable inner = parent.AddTable(1, 2, WordTableStyle.TableGrid);
        ConfigureNestedLayoutTable(inner, 6000, 3000, 3000);
        inner.WidthType = WordTableWidthUnit.Pct;
        inner.Width = innerWidthPercentage;
        foreach (WordTableCell cell in inner.Rows[0].Cells) {
            cell.WidthType = WordTableWidthUnit.Auto;
            cell.Width = 0;
        }
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[0].Paragraphs[0], "LEFT");
        ConfigureNestedLayoutParagraph(inner.Rows[0].Cells[1].Paragraphs[0], "RIGHT");
        string original = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(new WordToPdfOptions {
            IncludePageNumbers = false,
            PageSize = new PageSize(400D, 500D),
            Margins = PageMargins.Uniform(30D),
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        var before = Assert.Single(words, word => word.Text == "AHEAD");
        var left = Assert.Single(words, word => word.Text == "LEFT");
        var right = Assert.Single(words, word => word.Text == "RIGHT");
        Assert.Equal(left.BoundingBox.Bottom, right.BoundingBox.Bottom, precision: 2);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "NativeNestedTableLayoutApproximated");
        Assert.Equal(original, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        return (before.BoundingBox.Left, left.BoundingBox.Left, right.BoundingBox.Left - left.BoundingBox.Left);
    }

    private static void ConfigureNestedLayoutTable(WordTable table, int widthTwips, params int[] columns) {
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.Width = widthTwips;
        table.GridColumnWidth = columns.ToList();
        foreach (WordTableRow row in table.Rows) {
            for (int index = 0; index < row.Cells.Count; index++) {
                WordTableCell cell = row.Cells[index];
                cell.WidthType = WordTableWidthUnit.Dxa;
                cell.Width = columns[index];
                cell.MarginTopWidth = cell.MarginBottomWidth = cell.MarginLeftWidth = cell.MarginRightWidth = 0;
            }
        }
    }

    private static void ConfigureNestedLayoutParagraph(WordParagraph paragraph, string text) {
        paragraph.Text = text;
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = 10;
        paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
        paragraph.LineSpacing = 240;
        paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
        paragraph.KeepLinesTogetherOverride = paragraph.KeepWithNextOverride = false;
        paragraph.AvoidWidowAndOrphanOverride = false;
    }
}
