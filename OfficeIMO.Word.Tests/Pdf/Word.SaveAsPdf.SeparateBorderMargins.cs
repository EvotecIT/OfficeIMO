using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, 0, false)]
    [InlineData(0, 0, true)]
    [InlineData(0, 1, false)]
    [InlineData(0, 1, true)]
    [InlineData(0, 2, false)]
    [InlineData(0, 2, true)]
    [InlineData(120, 0, false)]
    [InlineData(120, 0, true)]
    [InlineData(120, 1, false)]
    [InlineData(120, 1, true)]
    [InlineData(120, 2, false)]
    [InlineData(120, 2, true)]
    public void SaveAsPdf_SeparateCellBordersRetainOwnTopMarginsWhileTextSharesTheRowFrame(
        short spacing, int heightRule, bool reverse) {
        using WordDocument source = WordDocument.Create();
        WordTable table = CreateBorderFrameControl(source, 1, 2, spacing);
        table.StyleDetails!.CellSpacing = spacing;
        WordTableRow row = table.Rows[0];
        if (heightRule == 1) row.Height = 2400;
        if (heightRule == 2) row.MinimumHeight = 2400;
        short margin = spacing == 0 ? (short)120 : (short)40;
        row.Cells[reverse ? 0 : 1].MarginTopWidth = margin;
        row.Cells[reverse ? 1 : 0].MarginBottomWidth = 180;

        using WordDocument imported = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdfBytes(BorderFramePdfOptions()));
        Assert.Equal(xml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var page = Assert.Single(pdf.GetPages());
        var cells = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(box => box.HasValue && box.Value.Width > 50D && box.Value.Width < 150D && box.Value.Height > 1D)
            .Select(box => box!.Value).OrderBy(box => box.Left).ToArray();
        Assert.Equal(2, cells.Length);
        Assert.Equal(reverse ? margin / 20D : -margin / 20D, cells[0].Top - cells[1].Top, 3);
        Assert.Equal(cells[0].Bottom, cells[1].Bottom, 3);
        var words = page.GetWords().ToArray();
        Assert.Equal(2, words.Length);
        Assert.Equal(words.Single(word => word.Text == "Frame0").Letters[0].StartBaseLine.Y,
            words.Single(word => word.Text == "Frame1").Letters[0].StartBaseLine.Y, 3);
    }
}
