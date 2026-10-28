using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, 1)]
    [InlineData(0, 2)]
    [InlineData(120, 1)]
    [InlineData(120, 2)]
    public void SaveAsPdf_RoundedSeparateCellsShareBorderPositionWithSquareNeighbors(short spacing, int columns) {
        using WordDocument source = WordDocument.Create();
        source.Sections[0].ColumnCount = columns;
        WordTable table = CreateBorderFrameControl(source, 1, 3, spacing);
        table._tableProperties!.TableStyle = null;
        table.StyleDetails!.CellSpacing = spacing;
        table.Rows[0].Height = global::OfficeIMO.OfficeLength.Points(120D);
        table.Rows[0].Cells[2].MarginTop = global::OfficeIMO.OfficeLength.Points(6D);
        WordPdfExportOptions options = BorderFramePdfOptions();
        options.PdfOptions = new OfficeIMO.Pdf.PdfCreateOptions {
            DefaultTableStyle = new OfficeIMO.Pdf.PdfTableStyle { CornerRadius = 8D }
        };

        using WordDocument imported = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(xml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var page = Assert.Single(pdf.GetPages());
        var cells = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(box => box.HasValue && box.Value.Width > 50D && box.Value.Width < 150D && box.Value.Height > 1D)
            .Select(box => box!.Value).GroupBy(box => Math.Round(box.Left, 3))
            .Select(group => group.First()).OrderBy(box => box.Left).ToArray();
        Assert.Equal(3, cells.Length);
        Assert.Equal(cells[0].Top, cells[1].Top, 3);
        Assert.Equal(6D, cells[2].Top - cells[0].Top, 3);
        Assert.All(cells, cell => Assert.Equal(cells[0].Bottom, cell.Bottom, 3));
        var words = page.GetWords().ToArray();
        Assert.Equal(3, words.Length);
        double baseline = words.Single(word => word.Text == "Frame0").Letters[0].StartBaseLine.Y;
        Assert.All(words, word => Assert.Equal(baseline, word.Letters[0].StartBaseLine.Y, 3));
    }

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
        if (heightRule == 1) row.Height = global::OfficeIMO.OfficeLength.Points(120D);
        if (heightRule == 2) row.MinimumHeight = global::OfficeIMO.OfficeLength.Points(120D);
        short margin = spacing == 0 ? (short)120 : (short)40;
        row.Cells[reverse ? 0 : 1].MarginTop = global::OfficeIMO.OfficeLength.Points((margin) / 20D);
        row.Cells[reverse ? 1 : 0].MarginBottom = global::OfficeIMO.OfficeLength.Points(9D);

        using WordDocument imported = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string xml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdf(BorderFramePdfOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
