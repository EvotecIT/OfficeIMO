using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfPageImageLayoutTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Svg)]
    public void WordPdfPageImagesUseThePdfPaginationAndPreserveEveryParagraph(OfficeImageExportFormat format) {
        using var word = WordDocument.Create();
        word.Sections[0].PageSettings.Width = 6000;
        word.Sections[0].PageSettings.Height = 6000;
        word.AddParagraph("Heading").SetStyle(WordParagraphStyles.Heading1);
        var table = word.AddTable(2, 1, WordTableStyle.TableGrid);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Table first row";
        table.Rows[1].Cells[0].Paragraphs[0].Text = "Table second row";
        for (int index = 0; index < 20; index++) word.AddParagraph("Paragraph" + index + " complete sentence");
        var conversion = word.ToPdfDocumentResult(new WordToPdfOptions {
            Margins = PageMargins.Uniform(30), ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        byte[] bytes = conversion.ToBytes();
        var pdf = PdfDocument.Load(bytes);
        var options = new PdfImageExportOptions { Scale = .5 };
        var images = conversion.ExportImages(format, options);
        Assert.True(images.Count > 1);
        Assert.Equal(pdf.Inspect().PageCount, images.Count);
        var direct = pdf.Render.ExportImages(format, options);
        for (int index = 0; index < images.Count; index++)
            Assert.Equal(direct[index].Bytes, images[index].Bytes);
        if (format == OfficeImageExportFormat.Svg) {
            string svg = string.Join("\n", images.Select(image => Encoding.UTF8.GetString(image.Bytes)));
            Assert.Contains("Heading", svg);
            Assert.Contains("Table first row", svg);
            for (int index = 0; index < 20; index++) Assert.Contains("Paragraph" + index, svg);
            Assert.DoesNotContain("…", svg);
        }
    }
}
