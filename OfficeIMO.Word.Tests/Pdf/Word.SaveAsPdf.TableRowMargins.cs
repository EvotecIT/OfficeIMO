using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, 0, 120D)]
    [InlineData(true, 0, 120D)]
    [InlineData(false, 1, 129D)]
    [InlineData(true, 1, 135D)]
    [InlineData(false, 2, 129D)]
    [InlineData(true, 2, 135D)]
    public void SaveAsPdf_UnspacedTableRowHeightIncludesEffectiveMargins(bool minimum, int marginMode, double advance) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(document, 2, 2, 0);
        foreach (WordTableRow row in table.Rows) {
            if (minimum) row.MinimumHeight = 2400;
            else row.Height = 2400;
            row.Cells[0].MarginTopWidth = marginMode == 0 ? (short)0 : (short)120;
            row.Cells[0].MarginBottomWidth = marginMode == 1 ? (short)180 : (short)0;
            row.Cells[1].MarginTopWidth = 0;
            row.Cells[1].MarginBottomWidth = marginMode == 2 ? (short)180 : (short)0;
        }

        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Equal(xml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Single(pdf.GetPages());
        var words = pdf.GetPage(1).GetWords().ToArray();
        Assert.Equal(4, words.Length);
        double first = words.Single(word => word.Text == "Frame0").Letters[0].StartBaseLine.Y;
        double next = words.Single(word => word.Text == "Frame2").Letters[0].StartBaseLine.Y;
        Assert.Equal(advance, first - next, 3);
    }
}
