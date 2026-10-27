using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointPdfTableFidelityTests {
    [Fact]
    public void DenseTableDefaultsGrowRowsAndKeepStatusOnOneLineAfterReopen() {
        // Minimal reproduction of p3-dense-layout: 76/56pt columns, 18/22pt
        // rows, empty tcPr/bodyPr and unsized runs using Medium Style 2 Accent 1.
        using var saved = new MemoryStream();
        using (var presentation = PowerPointDocument.Create()) {
            presentation.SlideSize.SetSizePoints(300, 200);
            presentation.SetThemeLatinFonts("Helvetica", "Helvetica");
            var table = presentation.AddSlide().AddTablePoints(2, 2, 20, 30, 132, 40);
            table.StyleId = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
            table.FirstRow = true;
            table.BandedRows = true;
            table.SetColumnWidthsPoints(76, 56);
            table.SetRowHeightsPoints(18, 22);
            table.GetCell(0, 0).Text = "Section";
            table.GetCell(0, 1).Text = "Status";
            table.GetCell(1, 0).Text = "Packed";
            table.GetCell(1, 1).Text = "Watch";
            table.GetCell(1, 1).HorizontalAlignment = PowerPointTextAlignment.Center;
            presentation.Save(saved);
        }
        saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        var conversion = reopened.ConvertToPdfWithEvidence();
        var item = TableItem(conversion.Value);
        var style = item.Block.Style!;
        Assert.Equal(7.2D, style.CellPaddingX);
        Assert.Equal(3.6D, style.CellPaddingY);
        Assert.True(item.Height > 40D);
        Assert.Equal(item.Height, style.RowMinHeights!.Sum(value => value!.Value), 5);
        Assert.True(item.Block.Cells[0][1].ShrinkTextToFit);
        Assert.Equal(PdfCore.PdfColor.White, Assert.Single(item.Block.Cells[0][1].Runs).Color);
        Assert.True(Assert.Single(item.Block.Cells[0][1].Runs).Bold);
        Assert.False(Assert.Single(item.Block.Cells[1][0].Runs).Bold);
        Assert.Equal("Helvetica", Assert.Single(item.Block.Cells[0][1].Runs).FontFamily);

        byte[] bytes = conversion.ToBytes();
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "table-cell-overflow");
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = parsed.GetPage(1);
        var lines = page.Letters.GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Select(line => string.Concat(line.OrderBy(letter => letter.StartBaseLine.X).Select(letter => letter.Value))).ToArray();
        Assert.Contains(lines, line => line.IndexOf("Status", StringComparison.Ordinal) >= 0);
        string text = string.Concat(page.Letters.Select(letter => letter.Value));
        Assert.Contains("Packed", text);
        Assert.Contains("Watch", text);
        Assert.InRange(page.Letters.Min(letter => letter.StartBaseLine.X), 27.1D, 27.3D);
        Assert.All(page.Letters, letter => Assert.True(letter.BoundingBox.Bottom >= page.Height - item.Y - item.Height - .1D));
    }

    [Theory]
    [InlineData(8)]
    [InlineData(24)]
    public void WrappedMergedCellGrowsItsRowsAndKeepsExplicitMarginsAndRunOverrides(int fontSize) {
        using var presentation = PowerPointDocument.Create();
        presentation.SlideSize.SetSizePoints(400, 300);
        presentation.SetThemeLatinFonts("Helvetica", "Helvetica");
        var table = presentation.AddSlide().AddTablePoints(2, 2, 20, 20, 240, 24);
        table.StyleId = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
        table.FirstRow = true;
        table.SetColumnWidthsPoints(120, 120);
        table.SetRowHeightsPoints(12, 12);
        table.MergeCells(0, 0, 1, 0);
        var cell = table.GetCell(0, 0);
        cell.Text = "Alpha beta gamma delta epsilon zeta";
        cell.FontSize = fontSize;
        cell.Bold = false;
        cell.Color = "FF0000";
        cell.PaddingLeftPoints = 0;
        cell.PaddingRightPoints = 9;
        cell.PaddingTopPoints = 8;
        cell.PaddingBottomPoints = 10;
        table.GetCell(0, 1).Text = "Header";
        table.GetCell(1, 1).Text = "Tail";

        var conversion = presentation.ConvertToPdfWithEvidence();
        var item = TableItem(conversion.Value);
        var mapped = item.Block.Cells[0][0];
        Assert.Equal(2, mapped.RowSpan);
        Assert.False(mapped.ShrinkTextToFit);
        var run = Assert.Single(mapped.Runs);
        Assert.False(run.Bold);
        Assert.Equal((double)fontSize, run.FontSize);
        Assert.Equal(PdfCore.PdfColor.FromRgb(255, 0, 0), run.Color);
        Assert.Equal(0D, item.Block.Style!.CellPaddings![(0, 0)].Left);
        Assert.Equal(10D, item.Block.Style.CellPaddings[(0, 0)].Bottom);
        Assert.True(item.Height > 24D);
        byte[] bytes = conversion.ToBytes();
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "table-cell-overflow");
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Contains("zeta", string.Concat(parsed.GetPage(1).Letters.Select(letter => letter.Value)));
    }

    [Fact]
    public void PlainTableHasNoPdfPresetFillAndLaterRunsDoNotInheritFirstRunsDirectFormatting() {
        using var presentation = PowerPointDocument.Create();
        presentation.SetThemeLatinFonts("Helvetica", "Helvetica");
        var table = presentation.AddSlide().AddTablePoints(1, 1, 20, 20, 240, 60);
        table.ApplyStyle(PowerPointTableStylePreset.Plain);
        table.StyleId = null;
        table.FirstRow = true;
        var cell = table.GetCell(0, 0);
        cell.Text = "Large";
        cell.Bold = true;
        cell.FontSize = 28;
        cell.Color = "FF0000";
        cell.AddRun(" regular");
        // AddRun copies the template for editing convenience; unspecify the
        // second run to reproduce genuine imported runs with no local style.
        cell.Cell.TextBody!.Elements<A.Paragraph>().Single().Elements<A.Run>().Last().RunProperties = new A.RunProperties();

        var conversion = presentation.ConvertToPdfWithEvidence();
        var item = TableItem(conversion.Value);
        var style = item.Block.Style!;
        Assert.Null(style.HeaderFill);
        Assert.Null(style.RowStripeFill);
        Assert.False(style.HeaderBold);
        Assert.True(item.Block.Cells[0][0].Runs[0].Bold);
        Assert.False(item.Block.Cells[0][0].Runs[1].Bold);
        Assert.NotEqual(28D, item.Block.Cells[0][0].Runs[1].FontSize);
        Assert.NotEqual(PdfCore.PdfColor.FromRgb(255, 0, 0), item.Block.Cells[0][0].Runs[1].Color);
        conversion.ToBytes();
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "table-cell-overflow");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyTablePreviewRetainsRowsThatGrowBeyondTheAuthoredFrame(bool onMaster) {
        using var presentation = PowerPointDocument.Create();
        presentation.SetThemeLatinFonts("Helvetica", "Helvetica");
        var table = presentation.AddSlide().AddTablePoints(1, 1, 20, 20, 120, 12);
        table.StyleId = null;
        var cell = table.GetCell(0, 0);
        cell.Text = "First\nSecond\nTail";
        cell.FontName = "Helvetica";
        cell.FontSize = 18;
        if (onMaster) {
            // Legacy export wraps master frames without attaching them to a
            // consuming slide. Theme tokens must resolve through that part.
            presentation.SetThemeColor(PowerPointThemeColor.Accent2, "112233");
            var master = presentation.OpenXmlDocument.PresentationPart!.SlideMasterParts.Single();
            master.SlideMaster!.ColorMap!.Accent1 = A.ColorSchemeIndexValues.Accent2;
            cell.FontName = "+mn-lt";
            var properties = cell.Cell.TextBody!.Elements<A.Paragraph>().Single().Elements<A.Run>().Single().RunProperties!;
            properties.RemoveAllChildren<A.SolidFill>();
            properties.Append(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }));
            var frame = (DocumentFormat.OpenXml.Presentation.GraphicFrame)table.Element.CloneNode(true);
            table.Element.Remove();
            master.SlideMaster.CommonSlideData!.ShapeTree!.Append(frame);
            table = new PowerPointTable(frame, master);
            Assert.Null(table.OwnerSlide);
        }

        var preview = PowerPointPdfConverterExtensions.CreateTableDrawingForExport(table);
        Assert.Equal(120D, preview.Width);
        Assert.True(preview.Height > 12D);
        Assert.NotEmpty(preview.Elements);
        // Read back the actual preview, including text below the original
        // frame, rather than only checking its enlarged model dimensions.
        var pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfCreateOptions {
            PageWidth = preview.Width,
            PageHeight = preview.Height,
            Margins = PdfCore.PdfPageMargins.Uniform(0)
        });
        pdf.Canvas(canvas => canvas.DrawingForClippedRendering(preview, 0, 0, preview.Width, preview.Height));
        byte[] bytes = pdf.ToBytes();
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(bytes);
        string text = string.Concat(parsed.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("First", text);
        Assert.Contains("Second", text);
        Assert.Contains("Tail", text);
        if (onMaster) Assert.Contains("0.067 0.133 0.2 rg", PdfOperatorSearchText.From(bytes), StringComparison.Ordinal);
    }

    private static PdfCore.PdfCanvasTableItem TableItem(PdfCore.PdfDocument document) =>
        Assert.Single(Assert.IsType<PdfCore.PdfCanvasBlock>(Assert.Single(document.Blocks)).Items.OfType<PdfCore.PdfCanvasTableItem>());
}
