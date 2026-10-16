using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeIMO.Tests;

public sealed class ExcelConditionalIconFidelityTests {
    [Theory]
    [InlineData(100, true)]
    [InlineData(50, true)]
    [InlineData(100, false)]
    [InlineData(50, false)]
    public void Symbols_RenderActualArtworkAndReserveTextInPdfAndPng(int percent, bool showValue) {
        using var document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Icons");
        sheet.SetColumnWidth(1, 20);
        sheet.SetMargins(0, 0, 0, 0);
        sheet.SetPageSetup(scale: (uint)percent);
        int[] values = { 123, 456, 789 };
        for (int row = 1; row <= 3; row++) {
            sheet.CellValue(row, 1, values[row - 1]);
            sheet.CellAlign(row, 1, ExcelHorizontalAlignment.Left);
            sheet.SetRowHeight(row, 24);
        }
        sheet.AddConditionalIconSet("A1:A3", ExcelIconSet.ThreeSymbols2, showValue, false,
            percentThresholds: null, numberThresholds: new[] { 0D, 300D, 600D });

        ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A3").CreateVisualSnapshot();
        Assert.Equal(new[] { OfficeConditionalIconKind.RedCross, OfficeConditionalIconKind.YellowExclamation, OfficeConditionalIconKind.GreenCheck },
            snapshot.ConditionalIcons.Select(icon => icon.Kind));

        byte[] pdfBytes = document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(360, 240), Margins = PdfCore.PageMargins.Uniform(0)
        });
        using var pdf = PdfPigDocument.Open(pdfBytes);
        double worksheetScale = percent / 100D;
        AssertPdfSymbolArtwork(pdfBytes, worksheetScale);
        var words = pdf.GetPage(1).GetWords().ToArray();
        foreach (int value in values) {
            var matching = words.Where(word => word.Text == value.ToString()).ToArray();
            if (!showValue) { Assert.Empty(matching); continue; }
            var word = Assert.Single(matching);
            // Nominal 3px inset plus 16px artwork, converted to PDF points.
            Assert.True(word.BoundingBox.Left > (3D + 16D) * 0.75D * worksheetScale + worksheetScale,
                "Left-aligned PDF text must begin after the icon and a positive gap.");
        }

        double imageScale = percent / 50D;
        byte[] png = sheet.Range("A1:A3").ToPng(new ExcelImageExportOptions { ShowGridlines = false, Scale = imageScale });
        Assert.True(OfficePngReader.TryDecode(png, out var image));
        AssertSymbolArtwork(image!, 32D * imageScale);
        foreach (var cell in snapshot.Cells) {
            var icon = snapshot.ConditionalIcons.Single(item => item.Row == cell.Row);
            var bounds = OfficeConditionalIconRenderer.GetCellBounds(cell.Width, cell.Height, imageScale);
            int rowTop = (int)(cell.Y * imageScale);
            int rowBottom = Math.Min(image!.Height, (int)((cell.Y + cell.Height) * imageScale));
            var blackX = new List<int>();
            for (int y = rowTop; y < rowBottom; y++) for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (pixel.R < 70 && pixel.G < 70 && pixel.B < 70) blackX.Add(x);
            }
            if (!icon.ShowValue) Assert.Empty(blackX);
            else {
                Assert.NotEmpty(blackX);
                Assert.True(blackX.Min() > bounds.X + bounds.Size, "PNG text must begin after the reserved icon box.");
            }
        }
    }

    [Fact]
    public void ImportedCustomNoIconsAndNumericThresholds_AgreeInPdfAndPng() {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            ExcelSheet sheet = document.AddWorksheet("Custom");
            sheet.CellValue(1, 1, -1);
            sheet.CellValue(2, 1, 0);
            sheet.CellValue(3, 1, 1);
            sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                Source = ExcelConditionalFormattingSource.Office2010Extension,
                Range = "A1:A3", Type = "IconSet", IconSet = "ThreeSymbols2", IconSetCustom = true,
                IconSetThresholds = new[] {
                    new ExcelConditionalIconSetThreshold { Type = "Percent", Value = "0" },
                    new ExcelConditionalIconSetThreshold { Type = "Number", Value = "0" },
                    new ExcelConditionalIconSetThreshold { Type = "Number", Value = "1" }
                },
                CustomIcons = Enumerable.Range(0, 3).Select(index => new ExcelConditionalFormatIcon { IconSet = "ThreeSymbols2", IconId = (uint)index }).ToArray()
            });
            // Producer-shaped x14 substitutions used by the audited workbook; authoring NoIcons is separate.
            foreach (var icon in sheet.WorksheetPart.Worksheet.Descendants<X14.ConditionalFormattingIcon>().Take(2))
                icon.SetAttribute(new OpenXmlAttribute("iconSet", null, "NoIcons"));
            document.Save();
        }
        using var loaded = ExcelDocument.Load(new MemoryStream(stream.ToArray()));
        ExcelSheet imported = loaded.Sheets.Single();
        var snapshot = imported.Range("A1:A3").CreateVisualSnapshot();
        var selected = Assert.Single(snapshot.ConditionalIcons);
        Assert.Equal(3, selected.Row);
        Assert.Equal(OfficeConditionalIconKind.GreenCheck, selected.Kind);
        byte[] pdf = loaded.ToPdfBytes(new ExcelToPdfOptions { IncludeSheetHeadings = false, HeaderRowCount = 0 });
        Assert.True(OfficePngReader.TryDecode(PdfCore.PdfPageImageRenderer.RenderPageAsPng(pdf, scale: 2D), out var image));
        Assert.True(FindColorBounds(image!, 2).Count > 0);
        Assert.Equal(0, FindColorBounds(image!, 0).Count);
        Assert.Equal(0, FindColorBounds(image!, 1).Count);
    }

    [Fact]
    public void IconBounds_StayAtLeftInWideCellsAndScaleWithTheCell() {
        var normal = OfficeConditionalIconRenderer.GetCellBounds(80, 32);
        var wide = OfficeConditionalIconRenderer.GetCellBounds(800, 32);
        Assert.Equal(3D, normal.X);
        Assert.Equal(normal, wide);
        var scaled = OfficeConditionalIconRenderer.GetCellBounds(800, 32, 0.5D);
        Assert.Equal((normal.X / 2D, normal.Y / 2D, normal.Size / 2D), scaled);
        var narrow = OfficeConditionalIconRenderer.GetCellBounds(10, 12);
        Assert.True(narrow.X + narrow.Size <= 10D);
        Assert.True(narrow.Y + narrow.Size <= 12D);
    }

    private static void AssertPdfSymbolArtwork(byte[] bytes, double scale) {
        // Inspect serialized paths with their CTM applied. Managed PDF raster projection
        // currently keeps unscaled stroke widths; Poppler separately verifies this fixture.
        var primitives = PdfCore.PdfReadDocument.Open(bytes).Pages[0].GetIdentityVisualPrimitives(scaleStrokeWidthWithTransform: true);
        var cross = primitives.Where(item => item.Kind == PdfCore.PdfPageVisualPrimitiveKind.Line &&
            item.StrokeColor is OfficeColor color && IsColor(color, 0)).ToArray();
        var check = primitives.Where(item => item.Kind == PdfCore.PdfPageVisualPrimitiveKind.Line &&
            item.StrokeColor is OfficeColor color && IsColor(color, 2)).ToArray();
        Assert.Equal(2, cross.Length);
        Assert.Equal(2, check.Length);
        // An X needs both diagonals, rather than a filled diamond or circular substitute.
        Assert.True((cross[0].X2 - cross[0].X1) * (cross[0].Y2 - cross[0].Y1) *
            (cross[1].X2 - cross[1].X1) * (cross[1].Y2 - cross[1].Y1) < 0D);
        Assert.Equal(check[0].X2, check[1].X1, 3);
        Assert.Equal(check[0].Y2, check[1].Y1, 3);
        Assert.True(Math.Abs(check[1].X2 - check[1].X1) > Math.Abs(check[0].X2 - check[0].X1));
        // The shared artwork uses a 14% stroke inside its nominal 12pt PDF box.
        Assert.All(cross.Concat(check), line => Assert.Equal(12D * 0.14D * scale, line.StrokeWidth, 3));

        var badge = Assert.Single(primitives.Where(item => item.FillColor is OfficeColor color &&
            IsColor(color, 1) && item.Y >= 24D * scale && item.Y < 48D * scale));
        Assert.Equal(12D * scale, badge.Width, 3);
        Assert.Equal(badge.Width, badge.Height, 3);
        var stem = Assert.Single(primitives.Where(item => item.Kind == PdfCore.PdfPageVisualPrimitiveKind.Line &&
            item.StrokeColor is OfficeColor color && color.R == 255 && color.G == 255 && color.B == 255));
        Assert.Equal(stem.X1, stem.X2, 3);
        Assert.True(Math.Abs(stem.Y2 - stem.Y1) > stem.StrokeWidth);
        var dot = Assert.Single(primitives.Where(item => item.FillColor is OfficeColor color &&
            color.R == 255 && color.G == 255 && color.B == 255 && (item.FillOpacity ?? 1D) == 1D));
        Assert.True(dot.Width > 0D && dot.Width < badge.Width / 4D);
        Assert.Equal(dot.Width, dot.Height, 3);
    }

    private static void AssertSymbolArtwork(OfficeRasterImage image, double rowHeight) {
        // Each authored 24pt row has one icon. Brown badge outlines and antialiased
        // red edges can satisfy another icon's color predicate, so inspect its own row.
        var cross = FindColorBounds(image, 0, 0D, rowHeight);
        var warning = FindColorBounds(image, 1, rowHeight, rowHeight * 2D);
        var check = FindColorBounds(image, 2, rowHeight * 2D, rowHeight * 3D);
        Assert.True(cross.Count > 4, "Expected a visible cross in the first row.");
        Assert.True(warning.Count > 4, "Expected a visible exclamation badge in the second row.");
        Assert.True(check.Count > 4, "Expected a visible check mark in the third row.");
        // Open strokes in the check/cross must not turn into filled circles or diamonds.
        Assert.True(check.Count < check.Area * 0.70D, "Check mark must have open background around its strokes.");
        Assert.True(cross.Count < cross.Area * 0.70D, "Cross must have open background around its strokes.");
        Assert.False(IsColor(image.GetPixel(cross.Left + 1, (cross.Top + cross.Bottom) / 2), 0));
        Assert.True(warning.Count > warning.Area * 0.45D, "Warning badge should surround its exclamation mark.");
        OfficeColor stem = image.GetPixel((warning.Left + warning.Right) / 2, warning.Top + (warning.Bottom - warning.Top) / 3);
        Assert.True(stem.R > 220 && stem.G > 220 && stem.B > 220, "Warning badge must contain a light exclamation stem.");
    }

    private static (int Left, int Top, int Right, int Bottom, int Count, int Area) FindColorBounds(
        OfficeRasterImage image, int color, double rowTop = 0D, double? rowBottom = null) {
        int left = image.Width, top = image.Height, right = -1, bottom = -1, count = 0;
        int firstY = Math.Max(0, (int)Math.Floor(rowTop));
        int lastY = Math.Min(image.Height, (int)Math.Ceiling(rowBottom ?? image.Height));
        for (int y = firstY; y < lastY; y++) for (int x = 0; x < image.Width; x++) {
            if (!IsColor(image.GetPixel(x, y), color)) continue;
            left = Math.Min(left, x); top = Math.Min(top, y); right = Math.Max(right, x); bottom = Math.Max(bottom, y); count++;
        }
        return (left, top, right, bottom, count, count == 0 ? 0 : (right - left + 1) * (bottom - top + 1));
    }

    private static bool IsColor(OfficeColor pixel, int color) => color switch {
        0 => pixel.R > 150 && pixel.G < 100 && pixel.B < 100,
        1 => pixel.R > 180 && pixel.G > 90 && pixel.G < 210 && pixel.B < 100,
        _ => pixel.G > 95 && pixel.G > pixel.R * 1.5D && pixel.G > pixel.B * 1.4D
    };
}
