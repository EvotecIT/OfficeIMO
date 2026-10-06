using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordFileFormat.Docx, 13, true)]
    [InlineData(WordFileFormat.Docx, 13, false)]
    [InlineData(WordFileFormat.Docx, 12, false)]
    [InlineData(WordFileFormat.Docx, 14, false)]
    [InlineData(WordFileFormat.Docx, 21, false)]
    [InlineData(WordFileFormat.Docx, 43, false)]
    [InlineData(WordFileFormat.Doc, 13, true)]
    [InlineData(WordFileFormat.Doc, 13, false)]
    [InlineData(WordFileFormat.Doc, 12, false)]
    [InlineData(WordFileFormat.Doc, 14, false)]
    [InlineData(WordFileFormat.Doc, 21, false)]
    [InlineData(WordFileFormat.Doc, 43, false)]
    public void SaveAsPdf_TableCellKeepRulesUseTheFullPageAfterAPartialColumnStart(
        WordFileFormat format, int lines, bool keepTogether) {
        byte[] source;
        using (WordDocument document = WordDocument.Create()) {
            document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2013;
            ConfigurePartialColumnParagraph(document.AddParagraph(string.Join("\n",
                Enumerable.Range(1, 12).Select(index => $"Before{index:D3}"))));
            WordSection columns = document.AddSection(WordSectionBreakType.Continuous);
            columns.ColumnCount=2; columns.ColumnsSpace=400;
            WordTable table = document.AddTable(1, 1);
            table.LayoutMode=WordTableLayoutMode.Fixed; table.Width=4000; table.WidthType=WordTableWidthUnit.Dxa;
            WordTableCell cell=table.Rows[0].Cells[0]; cell.Width=4000; cell.WidthType=WordTableWidthUnit.Dxa;
            cell.MarginTopWidth=0; cell.MarginBottomWidth=0; cell.MarginLeftWidth=0; cell.MarginRightWidth=0;
            WordParagraph paragraph=cell.Paragraphs[0];
            paragraph.Text=string.Join("\n", Enumerable.Range(1, lines).Select(index => $"Cell{index:D3}"));
            ConfigurePartialColumnParagraph(paragraph); paragraph.KeepLinesTogetherOverride=keepTogether;
            document.AddSection(WordSectionBreakType.Continuous).ColumnCount=1;
            var terminatorProperties = (DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties)columns._sectionProperties.Parent!;
            terminatorProperties.AddChild(new DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines {
                Line = "400", LineRule = DocumentFormat.OpenXml.Wordprocessing.LineSpacingRuleValues.Exact, After = "0"
            }, true);
            ConfigurePartialColumnParagraph(document.AddParagraph("AfterContinuous"));
            source=document.ToBytes(format);
        }
        string target=Path.Combine(_directoryWithFiles, $"PartialColumnTable-{format}-{lines}-{keepTogether}.pdf");
        using (var input=new MemoryStream(source))
        using (WordDocument document=WordDocument.Load(input)) document.SaveAsPdf(target, CellPaginationOptions());
        using var pdf=PdfPigDocument.Open(target);
        bool effectiveKeep=keepTogether && format == WordFileFormat.Docx;
        int finalPage=lines == 43 ? 3 : 2;
        Assert.Equal(finalPage, pdf.NumberOfPages);
        if (effectiveKeep) Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text.StartsWith("Cell", StringComparison.Ordinal));
        else Assert.InRange(FindWordStartX(pdf.GetPage(1), "Cell008"), 259.9, 260.1);
        int leftLast=format == WordFileFormat.Doc ? lines : effectiveKeep ? 13 : lines == 43 ? 42 : lines == 21 ? 15 : lines == 14 ? 11 : lines == 12 ? 10 : 11;
        Assert.InRange(FindWordStartX(pdf.GetPage(finalPage), $"Cell{leftLast:D3}"), 39.9, 40.1);
        if (format == WordFileFormat.Docx && !effectiveKeep)
            Assert.InRange(FindWordStartX(pdf.GetPage(finalPage), $"Cell{leftLast + 1:D3}"), 259.9, 260.1);
        Assert.InRange(FindWordStartX(pdf.GetPage(finalPage), "AfterContinuous"), 39.9, 40.1);
        int finalColumnLines = format == WordFileFormat.Doc ? lines - (finalPage == 3 ? 40 : 8) :
            effectiveKeep ? lines + 1 : (int)Math.Floor((lines - (finalPage == 3 ? 40 : 8)) / 2D) + 1;
        int tallestColumnLines = format == WordFileFormat.Doc || effectiveKeep ? finalColumnLines :
            Math.Max(leftLast - (finalPage == 3 ? 40 : 8), finalColumnLines);
        Assert.Equal(20D * tallestColumnLines,
            FindWordStartY(pdf.GetPage(finalPage), format == WordFileFormat.Doc ? $"Cell{(finalPage == 3 ? 41 : 9):D3}" :
                effectiveKeep ? "Cell001" : $"Cell{(finalPage == 3 ? 41 : 9):D3}") -
            FindWordStartY(pdf.GetPage(finalPage), "AfterContinuous"), 2);
        var words=Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).ToArray();
        foreach (int index in Enumerable.Range(1, lines)) Assert.Single(words, word => word.Text == $"Cell{index:D3}");
        foreach (int index in Enumerable.Range(1, 12)) Assert.Single(words, word => word.Text == $"Before{index:D3}");
        Assert.Single(words, word => word.Text == "AfterContinuous");
    }

    private static void ConfigurePartialColumnParagraph(WordParagraph paragraph) {
        paragraph.FontFamily="Arial"; paragraph.FontSize=12;
        paragraph.LineSpacing=400; paragraph.LineSpacingRule=WordLineSpacingRule.Exact;
        paragraph.LineSpacingBeforePoints=0; paragraph.LineSpacingAfterPoints=0;
        paragraph.KeepLinesTogetherOverride=false; paragraph.KeepWithNextOverride=false; paragraph.AvoidWidowAndOrphanOverride=false;
    }
}
