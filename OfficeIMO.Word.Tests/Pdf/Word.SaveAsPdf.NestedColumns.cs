using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2007, false)]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2010, false)]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2013, true)]
    [InlineData(WordFileFormat.Doc, WordCompatibilityMode.Word2013, false)]
    public void SaveAsPdf_BalancedKeepNextRetainsTheSourceModeTerminalGroupWithoutRetryingEmptyPages(
        WordFileFormat format, WordCompatibilityMode mode, bool keepsTerminalLine) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        document.Sections[0].ColumnCount = 2;
        document.Sections[0].ColumnsSpace = 400;
        WordParagraph leader = document.AddParagraph("Lead001\nLead002\nLead003");
        ConfigurePartialColumnParagraph(leader);
        leader.KeepWithNextOverride = true;
        WordParagraph body = document.AddParagraph("Body001\nBody002\nBody003\nBody004");
        ConfigurePartialColumnParagraph(body);
        body.KeepLinesTogetherOverride = true;
        document.AddSection(WordSectionBreakType.Continuous).ColumnCount = 1;
        ConfigurePartialColumnParagraph(document.AddParagraph("AfterContinuous"));
        string target = Path.Combine(_directoryWithFiles, $"BalancedKeepNext-{format}-{mode}.pdf");
        WordToPdfOptions options = CellPaginationOptions();
        options.PdfOptions = new OfficeIMO.Pdf.PdfOptions { MaxGeneratedPages = 4 };
        using var input = new MemoryStream(document.ToBytes(format));
        using (WordDocument reloaded = WordDocument.Load(input)) reloaded.SaveAsPdf(target, options);
        using var pdf = PdfPigDocument.Open(target);
        Assert.Equal(1, pdf.NumberOfPages);
        var page = pdf.GetPage(1);
        foreach (string marker in new[] { "Lead001", "Lead002", "Lead003", "Body001", "Body002", "Body003", "Body004", "AfterContinuous" })
            Assert.Single(page.GetWords(), word => word.Text == marker);
        Assert.InRange(FindWordStartX(page, "Lead001"), 39.9, 40.1);
        Assert.InRange(FindWordStartX(page, "Lead003"), keepsTerminalLine ? 259.9 : 39.9, keepsTerminalLine ? 260.1 : 40.1);
        Assert.InRange(FindWordStartX(page, "Body001"), 259.9, 260.1);
        Assert.InRange(FindWordStartX(page, "Body004"), 259.9, 260.1);
        Assert.Equal(80, FindWordStartY(page, "Body001") - FindWordStartY(page, "AfterContinuous"), 2);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2007, true)]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2010, true)]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2013, false)]
    [InlineData(WordFileFormat.Doc, WordCompatibilityMode.Word2013, true)]
    public void SaveAsPdf_PartialKeptParagraphUsesTheSourceColumnBalancingCompatibility(WordFileFormat format,
        WordCompatibilityMode mode, bool splitsBalancedColumns) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        ConfigurePartialColumnParagraph(document.AddParagraph(string.Join("\n", Enumerable.Range(1, 12).Select(index => $"Prelude{index:D3}"))));
        WordSection columns = document.AddSection(WordSectionBreakType.Continuous);
        columns.ColumnCount = 2; columns.ColumnsSpace = 400;
        WordParagraph paragraph = document.AddParagraph(string.Join("\n", Enumerable.Range(1, 13).Select(index => $"Marker{index:D3}")));
        ConfigurePartialColumnParagraph(paragraph); paragraph.KeepLinesTogetherOverride = true;
        document.AddSection(WordSectionBreakType.Continuous).ColumnCount = 1;
        ConfigurePartialColumnParagraph(document.AddParagraph("AfterContinuous"));
        string target = Path.Combine(_directoryWithFiles, $"PartialKeptMode-{format}-{mode}.pdf");
        using var input = new MemoryStream(document.ToBytes(format));
        using (WordDocument reloaded = WordDocument.Load(input)) reloaded.SaveAsPdf(target, CellPaginationOptions());
        using var pdf = PdfPigDocument.Open(target);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text.StartsWith("Marker", StringComparison.Ordinal));
        Assert.InRange(FindWordStartX(pdf.GetPage(2), "Marker007"), 39.9, 40.1);
        Assert.InRange(FindWordStartX(pdf.GetPage(2), "Marker008"), splitsBalancedColumns ? 259.9 : 39.9, splitsBalancedColumns ? 260.1 : 40.1);
        foreach (int index in Enumerable.Range(1, 13)) Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == $"Marker{index:D3}");
    }

    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void SaveAsPdf_NativeUnequalColumnsRetainAnOmittedGapAfterSaveReload(WordFileFormat format) {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].ColumnsSpace = 720;
        document.Sections[0].ColumnDefinitions = new[] { new WordSectionColumn(3000), new WordSectionColumn(4000, 0) };
        WordParagraph paragraph = document.AddParagraph();
        ConfigurePartialColumnParagraph(paragraph);
        paragraph.AddText("LeftColumn"); paragraph.AddBreak(WordBreakType.Column); paragraph.AddText("RightColumn");
        string before = Path.Combine(_directoryWithFiles, $"OmittedGapBefore-{format}.pdf");
        string after = Path.Combine(_directoryWithFiles, $"OmittedGapAfter-{format}.pdf");
        document.SaveAsPdf(before, CellPaginationOptions());
        using var source = new MemoryStream(document.ToBytes(format));
        using WordDocument reloaded = WordDocument.Load(source);
        reloaded.SaveAsPdf(after, CellPaginationOptions());
        using var beforePdf = PdfPigDocument.Open(before);
        using var afterPdf = PdfPigDocument.Open(after);
        double expected = FindWordStartX(beforePdf.GetPage(1), "RightColumn");
        Assert.InRange(expected, 189.9, 190.1);
        Assert.Equal(expected, FindWordStartX(afterPdf.GetPage(1), "RightColumn"), 2);
        Assert.Equal(1, afterPdf.NumberOfPages);
        Assert.Equal(40, FindWordStartX(afterPdf.GetPage(1), "LeftColumn"), 2);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void SaveAsPdf_AdjacentShadingCoversCollapsedSpacingWithoutMovingText(WordFileFormat format) {
        var positions = new List<Dictionary<string, (double X, double Y)>>();
        foreach (bool shaded in new[] { false, true }) {
            byte[] source;
            using (WordDocument document = WordDocument.Create()) {
                WordParagraph first = document.AddParagraph("FirstShaded");
                WordParagraph second = document.AddParagraph("SecondShaded");
                ConfigurePartialColumnParagraph(first); ConfigurePartialColumnParagraph(second);
                first.LineSpacingAfterPoints = 8; second.LineSpacingBeforePoints = 12; second.LineSpacingAfterPoints = 6;
                if (shaded) first.ShadingFillColorHex = second.ShadingFillColorHex = "00FFFF";
                ConfigurePartialColumnParagraph(document.AddParagraph("AfterShading"));
                source = document.ToBytes(format);
            }
            string target = Path.Combine(_directoryWithFiles, $"AdjacentShadingSpacing-{format}-{shaded}.pdf");
            using (var input = new MemoryStream(source))
            using (WordDocument document = WordDocument.Load(input)) document.SaveAsPdf(target, CellPaginationOptions());
            using var pdf = PdfPigDocument.Open(target);
            Assert.Equal(1, pdf.NumberOfPages);
            positions.Add(new[] { "FirstShaded", "SecondShaded", "AfterShading" }.ToDictionary(marker => marker,
                marker => (FindWordStartX(pdf.GetPage(1), marker), FindWordStartY(pdf.GetPage(1), marker))));
            if (shaded) {
                string raw = PdfOperatorSearchText.From(File.ReadAllBytes(target));
                var fills = ExtractFilledRectangles(raw, "0 1 1 rg").Where(fill => fill.Width > 400).OrderByDescending(fill => fill.Y).ToArray();
                Assert.Equal(2, fills.Length);
                Assert.InRange(Math.Abs(fills[0].Y - fills[1].Y - fills[1].Height), 0D, .01D);
            }
        }
        foreach (string marker in positions[0].Keys) Assert.Equal(positions[0][marker], positions[1][marker]);
        Assert.Equal(32, positions[1]["FirstShaded"].Y - positions[1]["SecondShaded"].Y, 2);
        Assert.Equal(26, positions[1]["SecondShaded"].Y - positions[1]["AfterShading"].Y, 2);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void SaveAsPdf_ShadingPreservesParagraphIndentsAndSpacingWithoutDoubleCountingPanelSpacing(WordFileFormat format) {
        var positions = new List<Dictionary<string, (double X, double Y)>>();
        foreach (bool shaded in new[] { false, true }) {
            byte[] source;
            using (WordDocument document = WordDocument.Create()) {
                ConfigurePartialColumnParagraph(document.AddParagraph("Before"));
                WordParagraph paragraph = document.AddParagraph("IndentedFirst\nIndentedSecond");
                ConfigurePartialColumnParagraph(paragraph);
                paragraph.IndentationBeforePoints = 12; paragraph.IndentationFirstLinePoints = 6;
                paragraph.LineSpacingBeforePoints = 12; paragraph.LineSpacingAfterPoints = 8;
                if (shaded) paragraph.ShadingFillColorHex = "00FFFF";
                ConfigurePartialColumnParagraph(document.AddParagraph("After"));
                source = document.ToBytes(format);
            }
            string target = Path.Combine(_directoryWithFiles, $"ShadedParagraphSpacing-{format}-{shaded}.pdf");
            using (var input = new MemoryStream(source))
            using (WordDocument document = WordDocument.Load(input)) document.SaveAsPdf(target, CellPaginationOptions());
            using var pdf = PdfPigDocument.Open(target);
            Assert.Equal(1, pdf.NumberOfPages);
            positions.Add(new[] { "Before", "IndentedFirst", "IndentedSecond", "After" }.ToDictionary(marker => marker,
                marker => (FindWordStartX(pdf.GetPage(1), marker), FindWordStartY(pdf.GetPage(1), marker))));
        }
        foreach (string marker in positions[0].Keys) {
            Assert.Equal(positions[0][marker].X, positions[1][marker].X, 2);
            Assert.Equal(positions[0][marker].Y, positions[1][marker].Y, 2);
        }
        Assert.Equal(58, positions[1]["IndentedFirst"].X, 2);
        Assert.Equal(52, positions[1]["IndentedSecond"].X, 2);
        Assert.Equal(20, positions[1]["IndentedFirst"].Y - positions[1]["IndentedSecond"].Y, 2);
        Assert.Equal(28, positions[1]["IndentedSecond"].Y - positions[1]["After"].Y, 2);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx)]
    [InlineData(WordFileFormat.Doc)]
    public void SaveAsPdf_ShadedKeepWithNextParagraphMovesWithItsFollowingBodyInColumns(WordFileFormat format) {
        byte[] source;
        using (WordDocument document = WordDocument.Create()) {
            document.Sections[0].ColumnCount = 2; document.Sections[0].ColumnsSpace = 400;
            ConfigurePartialColumnParagraph(document.AddParagraph(string.Join("\n", Enumerable.Range(1, 15).Select(index => $"Prelude{index:D3}"))));
            WordParagraph heading = document.AddParagraph("ShadedHeading");
            ConfigurePartialColumnParagraph(heading); heading.ShadingFillColorHex = "00FFFF"; heading.KeepWithNextOverride = true;
            ConfigurePartialColumnParagraph(document.AddParagraph("BodyFirst\nBodySecond\nBodyThird"));
            source = document.ToBytes(format);
        }
        string target = Path.Combine(_directoryWithFiles, $"ShadedKeepNext-{format}.pdf");
        using (var input = new MemoryStream(source))
        using (WordDocument document = WordDocument.Load(input)) document.SaveAsPdf(target, CellPaginationOptions());
        using var pdf = PdfPigDocument.Open(target);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(FindWordStartX(pdf.GetPage(1), "ShadedHeading"), 259.9, 260.1);
        Assert.InRange(FindWordStartX(pdf.GetPage(1), "BodyFirst"), 259.9, 260.1);
        Assert.Equal(20, FindWordStartY(pdf.GetPage(1), "ShadedHeading") - FindWordStartY(pdf.GetPage(1), "BodyFirst"), 2);
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, 21, false)]
    [InlineData(WordFileFormat.Docx, 43, false)]
    [InlineData(WordFileFormat.Docx, 21, true)]
    [InlineData(WordFileFormat.Docx, 43, true)]
    [InlineData(WordFileFormat.Doc, 21, false)]
    [InlineData(WordFileFormat.Doc, 43, false)]
    [InlineData(WordFileFormat.Doc, 21, true)]
    [InlineData(WordFileFormat.Doc, 43, true)]
    public void SaveAsPdf_DecoratedParagraphsPreserveExactSpacingAndBalancedContinuation(WordFileFormat format, int lines, bool border) {
        byte[] source;
        using (WordDocument document = WordDocument.Create()) {
            WordSection section = document.Sections[0]; section.ColumnCount = 2; section.ColumnsSpace = 400;
            WordParagraph paragraph = document.AddParagraph(string.Join("\n", Enumerable.Range(1, lines).Select(index => $"Nested{index:D3}")));
            ConfigurePartialColumnParagraph(paragraph);
            if (border) {
                WordParagraphBorders borders = paragraph.Borders;
                borders.LeftStyle = borders.RightStyle = borders.TopStyle = borders.BottomStyle = WordBorderStyle.Single;
                borders.LeftSize = borders.RightSize = borders.TopSize = borders.BottomSize = 4;
                borders.LeftSpace = borders.RightSpace = borders.TopSpace = borders.BottomSpace = 0;
            } else paragraph.ShadingFillColorHex = "00FFFF";
            document.AddSection(WordSectionBreakType.Continuous).ColumnCount = 1;
            ConfigurePartialColumnParagraph(document.AddParagraph("AfterContinuous"));
            source = document.ToBytes(format);
        }
        string target = Path.Combine(_directoryWithFiles, $"NestedColumns-{format}-{lines}-{border}.pdf");
        using (var input = new MemoryStream(source))
        using (WordDocument document = WordDocument.Load(input)) document.SaveAsPdf(target, CellPaginationOptions());
        using var pdf = PdfPigDocument.Open(target);
        int finalPage = lines == 21 ? 1 : 2;
        int leftLast = lines == 21 ? 11 : 38;
        Assert.Equal(finalPage, pdf.NumberOfPages);
        Assert.InRange(FindWordStartX(pdf.GetPage(finalPage), $"Nested{leftLast:D3}"), 39.9, 40.1);
        Assert.InRange(FindWordStartX(pdf.GetPage(finalPage), $"Nested{leftLast + 1:D3}"), 259.9, 260.1);
        int firstFinal = lines == 21 ? 1 : 33;
        Assert.Equal(20, FindWordStartY(pdf.GetPage(finalPage), $"Nested{firstFinal:D3}") -
            FindWordStartY(pdf.GetPage(finalPage), $"Nested{firstFinal + 1:D3}"), 2);
        Assert.Equal((leftLast - firstFinal + 1) * 20, FindWordStartY(pdf.GetPage(finalPage), $"Nested{firstFinal:D3}") -
            FindWordStartY(pdf.GetPage(finalPage), "AfterContinuous"), 2);
        var words = Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).ToArray();
        foreach (int index in Enumerable.Range(1, lines)) Assert.Single(words, word => word.Text == $"Nested{index:D3}");
        Assert.Single(words, word => word.Text == "AfterContinuous");
    }
}
