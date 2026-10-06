using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordFileFormat.Docx, 21, false)]
    [InlineData(WordFileFormat.Docx, 21, true)]
    [InlineData(WordFileFormat.Docx, 43, false)]
    [InlineData(WordFileFormat.Docx, 43, true)]
    [InlineData(WordFileFormat.Doc, 21, false)]
    [InlineData(WordFileFormat.Doc, 21, true)]
    [InlineData(WordFileFormat.Doc, 43, false)]
    [InlineData(WordFileFormat.Doc, 43, true)]
    [InlineData(WordFileFormat.Docx, 21, false, WordCompatibilityMode.Word2013)]
    [InlineData(WordFileFormat.Docx, 21, true, WordCompatibilityMode.Word2013)]
    [InlineData(WordFileFormat.Docx, 43, false, WordCompatibilityMode.Word2013)]
    [InlineData(WordFileFormat.Docx, 43, true, WordCompatibilityMode.Word2013)]
    public void SaveAsPdf_ContinuousColumns_HonorDocumentBalancingSuppression(
        WordFileFormat format, int lineCount, bool suppressBalancing,
        WordCompatibilityMode compatibilityMode = WordCompatibilityMode.Word2010) {
        string target = Path.Combine(_directoryWithFiles, $"ColumnBalance-{format}-{lineCount}-{suppressBalancing}-{compatibilityMode}.pdf");
        byte[] source;
        using (WordDocument document = WordDocument.Create()) {
            document.CompatibilitySettings.CompatibilityMode = compatibilityMode;
            // Use source XML so this also protects imports independently of the public setter.
            Settings settings = document._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
            Compatibility compatibility = settings.GetFirstChild<Compatibility>() ?? new Compatibility();
            if (compatibility.Parent == null) settings.AddChild(compatibility, true);
            compatibility.AddChild(new NoColumnBalance { Val = suppressBalancing }, true);
            document.Sections[0].ColumnCount = 2;
            document.Sections[0].ColumnsSpace = 400;
            AddParagraph(string.Join("\n", Enumerable.Range(1, lineCount).Select(index => $"BalanceLine{index:D3}")));
            document.AddSection(WordSectionBreakType.Continuous).ColumnCount = 1;
            AddParagraph("AfterContinuous");
            source = document.ToBytes(format);

            void AddParagraph(string text) {
                WordParagraph paragraph = document.AddParagraph(text);
                paragraph.FontFamily = "Arial";
                paragraph.FontSize = 12;
                paragraph.LineSpacing = 400;
                paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
                paragraph.LineSpacingBeforePoints = 0;
                paragraph.LineSpacingAfterPoints = 0;
                paragraph.AvoidWidowAndOrphanOverride = false;
            }
        }
        using (var stream = new MemoryStream(source))
        using (WordDocument loaded = WordDocument.Load(stream)) {
            Assert.Equal(suppressBalancing, loaded.CompatibilitySettings.DoNotBalanceTextColumns);
            loaded.SaveAsPdf(target, new WordToPdfOptions {
                IncludePageNumbers = false,
                PageSize = new PdfCore.PageSize(500, 400),
                Margins = PdfCore.PageMargins.Uniform(40)
            });
        }
        using var pdf = PdfPigDocument.Open(File.ReadAllBytes(target));
        bool effectiveSuppression = suppressBalancing && (format == WordFileFormat.Doc || compatibilityMode != WordCompatibilityMode.Word2013);
        Assert.Equal(lineCount == 21 && !effectiveSuppression ? 1 : 2, pdf.NumberOfPages);
        var lastContentPage = pdf.GetPage(lineCount == 21 ? 1 : 2);
        string lastMarker = $"BalanceLine{lineCount:D3}";
        double expectedLastX = effectiveSuppression && lineCount == 43 ? 40 : 260;
        Assert.InRange(FindWordStartX(lastContentPage, lastMarker), expectedLastX - .1, expectedLastX + .1);
        int firstRightLine = lineCount == 21 && !effectiveSuppression ? 12 : 17;
        Assert.InRange(FindWordStartX(pdf.GetPage(1), $"BalanceLine{firstRightLine:D3}"), 259.9, 260.1);
        var afterPage = pdf.GetPage(pdf.NumberOfPages);
        Assert.InRange(FindWordStartX(afterPage, "AfterContinuous"), 39.9, 40.1);
        string[] words = Enumerable.Range(1, pdf.NumberOfPages)
            .SelectMany(number => pdf.GetPage(number).GetWords()).Select(word => word.Text).ToArray();
        foreach (int index in Enumerable.Range(1, lineCount))
            Assert.Equal(1, words.Count(word => word == $"BalanceLine{index:D3}"));
        Assert.Equal(1, words.Count(word => word == "AfterContinuous"));
    }
}
