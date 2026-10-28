using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("text-box", false)]
    [InlineData("text-box", true)]
    [InlineData("numbered", false)]
    [InlineData("numbered", true)]
    public void SaveAsPdf_ParagraphBorderGroupDoesNotRemoveEdgesBesideAlternateRenderers(string route, bool specialFirst) {
        using WordDocument document = WordDocument.Create();
        WordParagraph special;
        if (route == "numbered") {
            WordList list = document.AddCustomList();
            var level = new WordListLevel(WordListLevelKind.DecimalDot) { LevelSuffix = WordListLevelSuffix.Tab };
            level.OpenXmlElement.NumberingSymbolRunProperties = new W.NumberingSymbolRunProperties(
                new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "24" });
            list.Numbering.AddLevel(level);
            special = list.AddItem("Special");
        } else {
            special = document.AddParagraph("Special");
            special.AddTextBox("Box", WordImageTextWrapping.Square);
        }
        WordParagraph ordinary = document.AddParagraph("Ordinary");
        if (!specialFirst) {
            ordinary._paragraph.Remove();
            special._paragraph.InsertBeforeSelf(ordinary._paragraph);
        }
        foreach (WordParagraph paragraph in new[] { ordinary, special }) {
            ConfigureAsymmetricBorderParagraph(paragraph);
            ConfigureMatchingParagraphBorders(paragraph, "box");
            paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Zero;
            paragraph.IndentationHanging = global::OfficeIMO.OfficeLength.Zero;
        }
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica"
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var page = pdf.GetPage(1);
        double baseline = Assert.Single(page.GetWords(), word => word.Text == "Ordinary").Letters[0].StartBaseLine.Y;
        var horizontal = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue && bounds.Value.Width > 300D && bounds.Value.Height < .01D)
            .Select(bounds => bounds!.Value.Bottom).ToArray();
        Assert.Contains(horizontal, y => y > baseline && y < baseline + 20D);
        Assert.Contains(horizontal, y => y < baseline && y > baseline - 20D);
    }

    public static IEnumerable<object[]> MatchingParagraphBorderCases() {
        foreach (WordFileFormat format in new[] { WordFileFormat.Docx, WordFileFormat.Doc }) {
            foreach (string shape in new[] { "box", "top", "bottom", "left", "top-left", "bottom-right" }) {
                yield return new object[] { format, shape };
            }
        }
    }

    [Theory]
    [MemberData(nameof(MatchingParagraphBorderCases))]
    public void SaveAsPdf_MatchingParagraphBordersReserveOnlyOuterSpacing(WordFileFormat format, string shape) {
        using WordDocument source = WordDocument.Create();
        WordParagraph[] paragraphs = new[] { "Before", "First", "Second", "Third", "After" }
            .Select(source.AddParagraph).ToArray();
        foreach (WordParagraph paragraph in paragraphs) ConfigureAsymmetricBorderParagraph(paragraph);
        foreach (WordParagraph paragraph in paragraphs.Skip(1).Take(3)) {
            ConfigureMatchingParagraphBorders(paragraph, shape);
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica",
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(360),OfficeIMO.OfficeLength.Points(300)), Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(1, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToArray();
        double[] y = new[] { "Before", "First", "Second", "Third", "After" }
            .Select(text => Assert.Single(words, word => word.Text == text).Letters[0].StartBaseLine.Y).ToArray();
        Assert.Equal(shape is "box" or "top" or "top-left" ? 20D : 16D, y[0] - y[1], 3);
        Assert.Equal(16D, y[1] - y[2], 3);
        Assert.Equal(16D, y[2] - y[3], 3);
        Assert.Equal(shape is "box" or "bottom" or "bottom-right" ? 20D : 16D, y[3] - y[4], 3);
    }

    [Theory]
    [InlineData(0U, 4U)]
    [InlineData(0U, 16U)]
    [InlineData(3U, 4U)]
    [InlineData(3U, 16U)]
    public void SaveAsPdf_SharedParagraphSeparatorUsesItsOwnSpacingAndThickness(uint space, uint size) {
        using WordDocument document = WordDocument.Create();
        foreach (string text in new[] { "First", "Second", "Third" }) {
            WordParagraph paragraph = document.AddParagraph(text);
            ConfigureAsymmetricBorderParagraph(paragraph);
            ConfigureMatchingParagraphBorders(paragraph, "box");
            paragraph._paragraph.ParagraphProperties!.ParagraphBorders!.BetweenBorder = new W.BetweenBorder {
                Val = W.BorderValues.Single, Color = "000000", Space = space, Size = size
            };
        }
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica"
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        double[] y = new[] { "First", "Second", "Third" }
            .Select(text => Assert.Single(words, word => word.Text == text).Letters[0].StartBaseLine.Y).ToArray();
        Assert.Equal(16D + 2D * space + size / 8D, y[0] - y[1], 3);
        Assert.Equal(16D + 2D * space + size / 8D, y[1] - y[2], 3);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, 0U, 3U)]
    [InlineData(WordFileFormat.Docx, 7U, 2U)]
    [InlineData(WordFileFormat.Docx, 31U, 0U)]
    [InlineData(WordFileFormat.Doc, 0U, 3U)]
    [InlineData(WordFileFormat.Doc, 7U, 2U)]
    [InlineData(WordFileFormat.Doc, 31U, 0U)]
    public void SaveAsPdf_SharedParagraphSeparatorPaintUsesTheFollowingParagraphSpace(
        WordFileFormat format, uint precedingSpace, uint followingSpace) {
        (double Baseline, double Separator) Render(uint middleSpace) {
            using WordDocument source = WordDocument.Create();
            var paragraphs = new[] { "First", "Second", "Third" }.Select(source.AddParagraph).ToArray();
            foreach (WordParagraph paragraph in paragraphs) {
                ConfigureAsymmetricBorderParagraph(paragraph);
                ConfigureMatchingParagraphBorders(paragraph, "box");
                paragraph._paragraph.ParagraphProperties!.ParagraphBorders!.BetweenBorder =
                    new W.BetweenBorder { Val = W.BorderValues.Single, Size = 8, Space = precedingSpace, Color = "000000" };
            }
            paragraphs[1]._paragraph.ParagraphProperties!.ParagraphBorders!.BetweenBorder!.Space = middleSpace;
            using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
            using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
                IncludePageNumbers = false, FontFamily = "Helvetica"
            }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
            var page = pdf.GetPage(1);
            double baseline = Assert.Single(page.GetWords(), word => word.Text == "First").Letters[0].StartBaseLine.Y;
            double[] strokes = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
                .Where(bounds => bounds.HasValue && bounds.Value.Width > 300D && bounds.Value.Height < .01D)
                .Select(bounds => bounds!.Value.Bottom).OrderByDescending(value => value).ToArray();
            Assert.Equal(4, strokes.Length);
            return (baseline, strokes[1]);
        }
        var reference = Render(precedingSpace);
        var changed = Render(followingSpace);
        Assert.Equal(reference.Baseline, changed.Baseline, 3);
        Assert.Equal((double)followingSpace - precedingSpace, reference.Separator - changed.Separator, 3);
    }

    [Theory]
    [InlineData("size", 23D, 24D)]
    [InlineData("color", 23D, 23D)]
    [InlineData("absent", 20D, 19D)]
    [InlineData("middle-only", 19D, 20D)]
    public void SaveAsPdf_ParagraphBorderSeparatorBelongsToThePrecedingParagraph(string variant, double firstAdvance, double secondAdvance) {
        using WordDocument document = WordDocument.Create();
        var paragraphs = new[] { "First", "Second", "Third" }.Select(document.AddParagraph).ToArray();
        foreach (WordParagraph paragraph in paragraphs) {
            ConfigureAsymmetricBorderParagraph(paragraph);
            ConfigureMatchingParagraphBorders(paragraph, "box");
            paragraph._paragraph.ParagraphProperties!.ParagraphBorders!.BetweenBorder =
                new W.BetweenBorder { Val = W.BorderValues.Single, Size = 8, Space = 3, Color = "000000" };
        }
        W.ParagraphBorders middle = paragraphs[1]._paragraph.ParagraphProperties!.ParagraphBorders!;
        if (variant == "size") middle.BetweenBorder!.Size = 16;
        if (variant == "color") middle.BetweenBorder!.Color = "FF0000";
        if (variant == "absent") middle.BetweenBorder!.Remove();
        if (variant == "middle-only") {
            paragraphs[0]._paragraph.ParagraphProperties!.ParagraphBorders!.BetweenBorder!.Remove();
            paragraphs[2]._paragraph.ParagraphProperties!.ParagraphBorders!.BetweenBorder!.Remove();
        }
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica"
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        double[] y = new[] { "First", "Second", "Third" }
            .Select(text => Assert.Single(words, word => word.Text == text).Letters[0].StartBaseLine.Y).ToArray();
        Assert.Equal(firstAdvance, y[0] - y[1], 3);
        Assert.Equal(secondAdvance, y[1] - y[2], 3);
    }

    private static void ConfigureMatchingParagraphBorders(WordParagraph paragraph, string shape) {
        paragraph.Borders.TopStyle = shape is "box" or "top" or "top-left" ? WordBorderStyle.Single : WordBorderStyle.Nil;
        paragraph.Borders.BottomStyle = shape is "box" or "bottom" or "bottom-right" ? WordBorderStyle.Single : WordBorderStyle.Nil;
        paragraph.Borders.LeftStyle = shape is "box" or "left" or "top-left" ? WordBorderStyle.Single : WordBorderStyle.Nil;
        paragraph.Borders.RightStyle = shape is "box" or "bottom-right" ? WordBorderStyle.Single : WordBorderStyle.Nil;
        paragraph.Borders.TopSize = paragraph.Borders.BottomSize = paragraph.Borders.LeftSize = paragraph.Borders.RightSize = 8;
        paragraph.Borders.TopSpace = paragraph.Borders.BottomSpace = paragraph.Borders.LeftSpace = paragraph.Borders.RightSpace = 3;
    }

    [Theory]
    [InlineData("box")]
    [InlineData("top-left")]
    public void SaveAsPdf_ParagraphBorderSeparatorReservesSpaceBeforeWidowSplit(string shape) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2013;
        ConfigureAsymmetricBorderParagraph(document.AddParagraph("Before"));
        foreach (string prefix in new[] { "First", "Second", "Third" }) {
            WordParagraph paragraph = document.AddParagraph(string.Join("\n", Enumerable.Range(1, 13).Select(i => $"{prefix}{i:D2}")));
            ConfigureAsymmetricBorderParagraph(paragraph);
            ConfigureMatchingParagraphBorders(paragraph, shape);
            paragraph._paragraph.ParagraphProperties!.ParagraphBorders!.BetweenBorder =
                new W.BetweenBorder { Val = W.BorderValues.Single, Size = 8, Space = 3, Color = "000000" };
        }
        ConfigureAsymmetricBorderParagraph(document.AddParagraph("After"));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica",
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(468),OfficeIMO.OfficeLength.Points(300)), Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(36))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(4, pdf.NumberOfPages);
        foreach (var item in new[] { ("First", 1), ("Second", 2), ("Third", 3) }) {
            for (int line = 1; line <= 13; line++) {
                int page = line <= 11 ? item.Item2 : item.Item2 + 1;
                Assert.Single(pdf.GetPage(page).GetWords(), word => word.Text == $"{item.Item1}{line:D2}");
            }
        }
        Assert.Single(pdf.GetPage(4).GetWords(), word => word.Text == "After");
    }

    [Theory]
    [InlineData("first-line", 16D)]
    [InlineData("hanging", 24D)]
    [InlineData("font", 16D)]
    [InlineData("auto-color", 24D)]
    [InlineData("frame", 24D)]
    [InlineData("left", 24D)]
    [InlineData("right", 24D)]
    [InlineData("between-space", 20D)]
    public void SaveAsPdf_ParagraphBorderGroupsHonorFrameBoundaries(string variant, double expectedAdvance) {
        using WordDocument document = WordDocument.Create();
        var paragraphs = new[] { "First", "Second", "Third" }.Select(document.AddParagraph).ToArray();
        foreach (WordParagraph paragraph in paragraphs) {
            ConfigureAsymmetricBorderParagraph(paragraph);
            ConfigureMatchingParagraphBorders(paragraph, "box");
            foreach (W.BorderType border in paragraph._paragraph.ParagraphProperties!.ParagraphBorders!.ChildElements.OfType<W.BorderType>())
                border.Color = "000000";
            if (variant == "between-space") paragraph._paragraph.ParagraphProperties.ParagraphBorders.BetweenBorder =
                new W.BetweenBorder { Val = W.BorderValues.Single, Size = 8, Space = 3, Color = "000000" };
        }
        WordParagraph middle = paragraphs[1];
        switch (variant) {
            case "first-line": middle.IndentationFirstLine = global::OfficeIMO.OfficeLength.Points(18D); break;
            case "hanging": middle.IndentationHanging = global::OfficeIMO.OfficeLength.Points(18D); break;
            case "font": middle.FontFamily = "Times New Roman"; break;
            case "left": middle.IndentationBefore = global::OfficeIMO.OfficeLength.Points(18D); break;
            case "right": middle.IndentationAfter = global::OfficeIMO.OfficeLength.Points(18D); break;
            case "between-space": middle._paragraph.ParagraphProperties!.ParagraphBorders!.BetweenBorder!.Space = 0; break;
            case "frame":
            case "auto-color":
                foreach (W.BorderType border in middle._paragraph.ParagraphProperties!.ParagraphBorders!.ChildElements.OfType<W.BorderType>()) {
                    if (variant == "frame") border.Frame = true;
                    else border.Color = "auto";
                }
                break;
        }
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, FontFamily = "Helvetica"
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = pdf.GetPage(1).GetWords().ToArray();
        double[] y = new[] { "First", "Second", "Third" }
            .Select(text => Assert.Single(words, word => word.Text == text).Letters[0].StartBaseLine.Y).ToArray();
        Assert.Equal(expectedAdvance, y[0] - y[1], 3);
        Assert.Equal(expectedAdvance, y[1] - y[2], 3);
    }
}
