using OfficeIMO.Pdf.Word;
using OfficeIMO.Word;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfSemanticImport_ModernTableRespectsTheRepeatingHeaderOption(bool repeat) {
        byte[] pdf = PdfCore.PdfDocument.Create()
            .Table(new[] {
                new[] { "Code", "Value" },
                new[] { "A", "1" },
                new[] { "B", "2" }
            }, style: new PdfCore.PdfTableStyle { HeaderRowCount = 1 })
            .ToBytes();
        using var imported = LoadSemanticPdf(pdf).ConvertToWordWithEvidence(new WordPdfImportOptions {
            TableStyle = WordTableStyle.OfficeIMOModern, RepeatHeaderRows = repeat
        }).Value;
        var table = Assert.Single(imported.Tables);
        Assert.Equal(repeat, table.RepeatHeaderRowAtTheTopOfEachPage);
        Assert.Equal("Code", table.Rows[0].Cells[0].Paragraphs[0].Text);
        using var saved = WordDocument.Load(new System.IO.MemoryStream(imported.ToBytes()));
        Assert.Equal(repeat, Assert.Single(saved.Tables).RepeatHeaderRowAtTheTopOfEachPage);
    }
}
