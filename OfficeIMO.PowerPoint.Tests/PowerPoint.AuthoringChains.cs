using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.Drawing;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointAuthoringChainsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WholeTextFormattingCoversRunsAndLineBreaksAndClearsNullableOverrides(bool cellTarget) {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        if (cellTarget) {
            var cell = slide.AddTable(1, 1).GetCell(0, 0);
            cell.Text = "First\vSecond";
            cell.AddParagraph("Third");
            Assert.Same(cell, cell.SetBold().SetItalic().SetFontFamily("Aptos")
                .SetFontSize(OfficeLength.Points(17)).SetColor(OfficeColor.Accent4));
        } else {
            var box = slide.AddTextBox("First\vSecond");
            box.AddParagraph("Third");
            Assert.Same(box, box.SetBold().SetItalic().SetFontFamily("Aptos")
                .SetFontSize(OfficeLength.Points(17)).SetColor(OfficeColor.Accent4));
        }
        byte[] bytes = document.ToBytes();
        using (var package = PresentationDocument.Open(new MemoryStream(bytes), false)) {
            var part = package.PresentationPart!.SlideParts.Single();
            var properties = part.Slide!.Descendants<A.RunProperties>().ToArray();
            Assert.True(properties.Length >= 3);
            Assert.All(properties, native => {
                Assert.True(native.Bold!.Value && native.Italic!.Value);
                Assert.Equal(1700, native.FontSize!.Value);
                Assert.Equal("Aptos", native.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
                Assert.Equal("accent4", native.GetFirstChild<A.SolidFill>()!.SchemeColor!.Val!.InnerText);
            });
            Assert.NotEmpty(part.Slide!.Descendants<A.Break>());
        }
        using var loaded = PowerPointDocument.Load(new MemoryStream(bytes));
        if (cellTarget) loaded.Slides[0].Tables.First().GetCell(0, 0).SetFontFamily(null).SetFontSize(null).SetColor(null);
        else loaded.Slides[0].TextBoxes.First().SetFontFamily(null).SetFontSize(null).SetColor(null);
        using var cleared = PresentationDocument.Open(new MemoryStream(loaded.ToBytes()), false);
        Assert.All(cleared.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.RunProperties>(), native => {
            Assert.Null(native.FontSize);
            Assert.Null(native.GetFirstChild<A.LatinFont>());
            Assert.Null(native.GetFirstChild<A.SolidFill>());
            Assert.True(native.Bold!.Value && native.Italic!.Value);
        });
    }

    [Fact]
    public void TypedShapeAndChartColorsRetainThemeSlotsAlphaAndFillChoices() {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var shape = slide.AddTextBox("Styled");
        shape.FillColor = OfficeColor.Accent1;
        shape.OutlineColor = OfficeColor.Parse("33669980");
        var chart = slide.AddChart(OfficeChartType.Line,
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Values", new[] { 1d, 2d }) }));
        chart.SetDataLabelShapeStyle(noFill: true, noLine: true)
            .SetDataLabelShapeStyle(OfficeColor.Accent2, OfficeColor.Accent3)
            .SetDataLabelTextStyle(italic: true, color: OfficeColor.Accent4, fontFamily: "Aptos")
            .SetSeriesMarker(0, OfficeIMO.Drawing.OfficeChartMarkerShape.Circle, fillColor: OfficeColor.Accent5)
            .SetSeriesFillColor(0, OfficeColor.Accent6);
        using var package = PresentationDocument.Open(new MemoryStream(document.ToBytes()), false);
        var slidePart = package.PresentationPart!.SlideParts.Single();
        var nativeShape = slidePart.Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().Single();
        Assert.Equal("accent1", nativeShape.ShapeProperties!.GetFirstChild<A.SolidFill>()!.SchemeColor!.Val!.InnerText);
        var outline = nativeShape.ShapeProperties!.GetFirstChild<A.Outline>()!;
        Assert.Equal(50196, outline.Descendants<A.Alpha>().Single().Val!.Value);
        var labels = slidePart.ChartParts.Single().ChartSpace!.Descendants<C.DataLabels>().First();
        var labelShape = labels.GetFirstChild<C.ChartShapeProperties>()!;
        Assert.Null(labelShape.GetFirstChild<A.NoFill>());
        Assert.Null(labelShape.GetFirstChild<A.Outline>()!.GetFirstChild<A.NoFill>());
        Assert.Equal("accent2", labelShape.GetFirstChild<A.SolidFill>()!.SchemeColor!.Val!.InnerText);
        using var restored = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(OfficeColor.Accent1, restored.Slides[0].TextBoxes.First().FillColor);
        Assert.Equal(OfficeColor.Parse("33669980"), restored.Slides[0].TextBoxes.First().OutlineColor);
    }

    [Fact]
    public void TableAndGeometryChainsKeepGridSynchronizationAndTheResizeAnchor() {
        using var document = PowerPointDocument.Create();
        var table = document.AddSlide().AddTable(2, 2);
        Assert.Same(table, table.SetColumnWidths(OfficeLength.Points(40), OfficeLength.Points(60))
            .SetRowHeights(OfficeLength.Points(20), OfficeLength.Points(30))
            .SetCellPadding(OfficeLength.Points(2), null, null, null));
        Assert.Equal(OfficeLength.Points(100), table.Width);
        Assert.Equal(OfficeLength.Points(50), table.Height);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(2, table.Columns.Count);
        Assert.Equal(2, table.RowCount);
        Assert.Equal(2, table.ColumnCount);
        var inserted = table.AddRow(1);
        Assert.Equal(3, table.RowCount);
        Assert.Equal(2, inserted.Cells.Count);
        var shape = table.SetPosition(OfficeLength.Points(10), OfficeLength.Points(20))
            .SetSize(OfficeLength.Points(100), OfficeLength.Points(50));
        var centerX = shape.CenterX;
        var centerY = shape.CenterY;
        Assert.Same(table, shape.Resize(OfficeLength.Points(80), OfficeLength.Points(40), PowerPointShapeAnchor.Center)
            .Scale(2, PowerPointShapeAnchor.Center).MoveBy(OfficeLength.Points(3), OfficeLength.Points(-2)));
        Assert.Equal(centerX + OfficeLength.Points(3), shape.CenterX);
        Assert.Equal(centerY - OfficeLength.Points(2), shape.CenterY);
        using var loaded = PowerPointDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(3, loaded.Slides[0].Tables.First().RowCount);
    }
}
