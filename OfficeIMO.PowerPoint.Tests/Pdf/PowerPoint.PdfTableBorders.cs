using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointPdfTableBorderTests {
    [Fact]
    public void Resolved_cell_edges_replace_the_PDF_preset_grid_after_saved_PPTX_reopen() {
        using var saved = new MemoryStream();
        using (var presentation = PowerPointDocument.Create()) {
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(260),OfficeIMO.OfficeLength.Points(180));
            var table = presentation.AddSlide().AddTable(2,2,global::OfficeIMO.OfficeLength.Points(30),global::OfficeIMO.OfficeLength.Points(34),global::OfficeIMO.OfficeLength.Points(160),global::OfficeIMO.OfficeLength.Points(80));
            table.ApplyStyle(PowerPointTableStylePreset.Plain);
            for (int row = 0; row < 2; row++) for (int column = 0; column < 2; column++) {
                var cell = table.GetCell(row, column);
                cell.Text = row == 0 ? "Plain" : "Body";
                cell.Color = "000000";
                cell.FillColor = "FFFFFF";
                cell.SetBorders(PowerPointTableCellBorders.All, "000000", global::OfficeIMO.OfficeLength.Points(2));
            }
            table.GetCell(0, 0).HideBorders(PowerPointTableCellBorders.Right);
            table.GetCell(0, 1).HideBorders(PowerPointTableCellBorders.Left);
            table.GetCell(0, 0).SetBorders(PowerPointTableCellBorders.Bottom, "FF3333", global::OfficeIMO.OfficeLength.Points(3));
            table.GetCell(1, 0).SetBorders(PowerPointTableCellBorders.Top, "FF3333", global::OfficeIMO.OfficeLength.Points(3));
            presentation.Save(saved);
        }
        saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        Assert.Empty(reopened.ValidateDocument());
        var document = reopened.ToPdf();
        var canvas = Assert.IsType<PdfCore.PdfCanvasBlock>(Assert.Single(document.Blocks));
        var style = Assert.Single(canvas.Items.OfType<PdfCore.PdfCanvasTableItem>()).Block.Style!;
        Assert.Null(style.BorderColor);
        Assert.Null(style.CellBorders![(0, 0)].RightBorder!.Color);
        Assert.Equal(0, style.CellBorders[(0, 0)].RightBorder!.Width);
        Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(255, 51, 51), style.CellBorders[(0, 0)].BottomBorder!.Color);
        Assert.Equal(3, style.CellBorders[(0, 0)].BottomBorder!.Width);
        string operators = PdfOperatorSearchText.From(document.ToBytes(new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains("0 0 0 RG", operators, StringComparison.Ordinal);
        Assert.Contains("1 0.2 0.2 RG", operators, StringComparison.Ordinal);
        Assert.DoesNotContain("0.824 0.855 0.886 RG", operators, StringComparison.Ordinal);
        Assert.DoesNotContain("110 146 m\n110 106 l", operators, StringComparison.Ordinal);
    }
}
