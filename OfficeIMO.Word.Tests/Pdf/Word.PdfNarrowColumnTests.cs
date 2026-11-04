using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void PdfExport_WordAuthoredTablesKeepShortWordsWholeBesideSpacerColumns() {
        string source = Path.Combine(_directoryDocuments, "sample1.docx");
        string output = Path.Combine(_directoryWithFiles, "Sample1NarrowColumns.pdf");
        using (WordDocument document = WordDocument.Load(source)) document.SaveAsPdf(output);

        using var pdf = PdfPigDocument.Open(output);
        string[] words = Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).Select(word => word.Text).ToArray();
        Assert.Contains("Sun", words);
        Assert.Contains("Mon", words);
        Assert.Contains("15", words);
        Assert.Contains("31", words);
        Assert.DoesNotContain("Su", words);

        var page3 = pdf.GetPage(3).GetWords().ToArray();
        var point = page3.Where(word => word.Text == "Point").OrderBy(word => word.BoundingBox.Left).Last();
        Assert.Contains(page3, word => word.Text == "E" && Math.Abs(word.BoundingBox.Bottom - point.BoundingBox.Bottom) < 1D);
    }
}
