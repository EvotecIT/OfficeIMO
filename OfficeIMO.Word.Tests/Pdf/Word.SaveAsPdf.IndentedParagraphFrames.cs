using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordFileFormat.Docx, 36D, 0D)]
    [InlineData(WordFileFormat.Doc, 36D, 0D)]
    [InlineData(WordFileFormat.Docx, 36D, -12D)]
    [InlineData(WordFileFormat.Docx, -12D, 0D)]
    public void SaveAsPdf_ParagraphFrameFollowsIndentsWithoutMovingTextTwice(
        WordFileFormat format, double left, double firstLine) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph(string.Join("\n", Enumerable.Range(1, 24).Select(i => $"Marker{i:D3}")));
        ConfigureAsymmetricBorderParagraph(paragraph);
        paragraph.IndentationBeforePoints = left;
        paragraph.IndentationAfterPoints = 18D;
        if (firstLine < 0D) paragraph.IndentationHangingPoints = -firstLine;
        paragraph.Borders.LeftStyle = paragraph.Borders.RightStyle = WordBorderStyle.Single;
        paragraph.Borders.LeftSize = paragraph.Borders.RightSize = 8;
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes(format)));
        using var model = document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new OfficeIMO.Pdf.PdfPageSize(360, 260),
            Margins = OfficeIMO.Pdf.PdfPageMargins.Uniform(30),
            FontFamily = "Helvetica"
        });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(model.ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));

        Assert.True(pdf.NumberOfPages > 1);
        int count = 0;
        foreach (var page in pdf.GetPages()) {
            var vertical = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
                .Where(bounds => bounds.HasValue && bounds.Value.Width < .01D && bounds.Value.Height > 1D)
                .Select(bounds => bounds!.Value.Left).ToArray();
            Assert.Contains(vertical, x => Math.Abs(x - (30D + left + Math.Min(firstLine, 0D) - 2D)) < .01D);
            Assert.Contains(vertical, x => Math.Abs(x - 314D) < .01D);
            foreach (var word in page.GetWords()) {
                count++;
                Assert.Equal($"Marker{count:D3}", word.Text);
                Assert.Equal(30D + left + (count == 1 ? firstLine : 0D), word.Letters[0].StartBaseLine.X, 2);
            }
        }
        Assert.Equal(24, count);
    }
}
