using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2007, 21)]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2007, 43)]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2010, 21)]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2010, 43)]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2013, 21)]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2013, 43)]
    [InlineData(WordFileFormat.Doc, WordCompatibilityMode.Word2013, 21)]
    [InlineData(WordFileFormat.Doc, WordCompatibilityMode.Word2013, 43)]
    public void SaveAsPdf_ContinuousColumns_BalanceTableRowLinesForModernDocxLayout(
        WordFileFormat format, WordCompatibilityMode mode, int lines) {
        byte[] source;
        using (WordDocument document = WordDocument.Create()) {
            document.CompatibilitySettings.CompatibilityMode = mode;
            WordTable table = CreatePaginationTable(document, 1);
            table.Rows[0].Cells[0].Paragraphs[0].Text = string.Join("\n",
                Enumerable.Range(1, lines).Select(index => $"Cell{index:D3}"));
            document.AddSection(WordSectionBreakType.Continuous).ColumnCount = 1;
            WordParagraph after = document.AddParagraph("AfterContinuous");
            after.FontFamily = "Arial"; after.FontSize = 12;
            after.LineSpacing = 400; after.LineSpacingRule = WordLineSpacingRule.Exact;
            after.LineSpacingBeforePoints = 0; after.LineSpacingAfterPoints = 0;
            after.KeepLinesTogetherOverride = false; after.KeepWithNextOverride = false;
            after.AvoidWidowAndOrphanOverride = false;
            source = document.ToBytes(format);
        }
        string target = Path.Combine(_directoryWithFiles, $"TableRowBalance-{format}-{mode}-{lines}.pdf");
        using (var input = new MemoryStream(source))
        using (WordDocument document = WordDocument.Load(input)) document.SaveAsPdf(target, CellPaginationOptions());
        using var pdf = PdfPigDocument.Open(target);
        bool modern = format == WordFileFormat.Docx && mode == WordCompatibilityMode.Word2013;
        Assert.Equal(lines == 21 && modern ? 1 : 2, pdf.NumberOfPages);
        int contentPage = lines == 21 ? 1 : 2;
        int leftLast = lines == 21 ? (modern ? 11 : 16) : (modern ? 38 : 43);
        Assert.InRange(FindWordStartX(pdf.GetPage(contentPage), $"Cell{leftLast:D3}"), 39.9, 40.1);
        if (modern || lines == 21)
            Assert.InRange(FindWordStartX(pdf.GetPage(contentPage), $"Cell{leftLast + 1:D3}"), 259.9, 260.1);
        Assert.InRange(FindWordStartX(pdf.GetPage(pdf.NumberOfPages), "AfterContinuous"), 39.9, 40.1);
        var words = Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).ToArray();
        foreach (int index in Enumerable.Range(1, lines)) Assert.Single(words, word => word.Text == $"Cell{index:D3}");
        Assert.Single(words, word => word.Text == "AfterContinuous");
    }
}
