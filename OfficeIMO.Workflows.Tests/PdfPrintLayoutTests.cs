using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfPrintLayoutTests {
    [Theory]
    [InlineData(6, false, 2, 3)]
    [InlineData(6, true, 3, 2)]
    [InlineData(9, false, 3, 3)]
    [InlineData(9, true, 3, 3)]
    public void LargerNUpPreservesPageOrderAndKeepsSlotsInsideAsymmetricMargins(int count, bool landscape, int columns, int rows) {
        PdfDocument document = Pages(count + 1);
        PdfPrintPlan plan = PdfPrintPlanner.Create(document, new() {
            InputPath = "snapshot.pdf", Pages = string.Join(',', Enumerable.Range(1, count + 1).Reverse()),
            PagesPerSheet = count, PaperSize = new PageSize(300, 500),
            Orientation = landscape ? PdfPrintOrientation.Landscape : PdfPrintOrientation.Portrait,
            MarginLeft = 10, MarginTop = 20, MarginRight = 30, MarginBottom = 40
        });
        Assert.Equal(2, plan.Sheets.Count);
        PdfPrintSheet sheet = plan.Sheets[0];
        Assert.Equal(Enumerable.Range(2, count).Reverse(), sheet.Placements.Select(p => p.PageNumber));
        Assert.Equal(columns, sheet.Placements.Select(p => p.SlotX).Distinct().Count());
        Assert.Equal(rows, sheet.Placements.Select(p => p.SlotY).Distinct().Count());
        Assert.Equal(10, sheet.Placements.Min(p => p.SlotX));
        Assert.Equal(20, sheet.Placements.Min(p => p.SlotY));
        Assert.Equal(sheet.PaperSize.Width - 30, sheet.Placements.Max(p => p.SlotX + p.SlotWidth), 6);
        Assert.Equal(sheet.PaperSize.Height - 40, sheet.Placements.Max(p => p.SlotY + p.SlotHeight), 6);
        Assert.All(sheet.Placements, p => {
            Assert.False(p.IsClipped);
            Assert.InRange(p.X, p.SlotX, p.SlotX + p.SlotWidth - p.Width + 0.000001);
            Assert.InRange(p.Y, p.SlotY, p.SlotY + p.SlotHeight - p.Height + 0.000001);
        });
        Assert.Equal(1, Assert.Single(plan.Sheets[1].Placements).PageNumber);
    }

    [Theory]
    [InlineData(PdfPrintAlignment.TopLeft, 0, 0)]
    [InlineData(PdfPrintAlignment.Top, 0.5, 0)]
    [InlineData(PdfPrintAlignment.TopRight, 1, 0)]
    [InlineData(PdfPrintAlignment.Left, 0, 0.5)]
    [InlineData(PdfPrintAlignment.Center, 0.5, 0.5)]
    [InlineData(PdfPrintAlignment.Right, 1, 0.5)]
    [InlineData(PdfPrintAlignment.BottomLeft, 0, 1)]
    [InlineData(PdfPrintAlignment.Bottom, 0.5, 1)]
    [InlineData(PdfPrintAlignment.BottomRight, 1, 1)]
    public void CustomPhysicalScaleUsesRequestedAnchorForBothWhitespaceAndOverflow(PdfPrintAlignment alignment, double horizontal, double vertical) {
        foreach (double percent in new[] { 50D, 250D }) {
            PdfPrintPlacement p = Assert.Single(Assert.Single(PdfPrintPlanner.Create(Pages(1), new() {
                InputPath = "snapshot.pdf", PaperSize = new PageSize(150, 180), Orientation = PdfPrintOrientation.Portrait,
                Margin = 10, ScaleMode = PdfPrintScaleMode.Custom, CustomScalePercent = percent, Alignment = alignment
            }).Sheets).Placements);
            Assert.Equal(percent / 100D, p.Scale);
            Assert.Equal(100 * percent / 100D, p.Width);
            Assert.Equal(120 * percent / 100D, p.Height);
            Assert.Equal(10 + (130 - p.Width) * horizontal, p.X, 6);
            Assert.Equal(10 + (160 - p.Height) * vertical, p.Y, 6);
            Assert.Equal(percent > 100, p.IsClipped);
        }
    }

    [Theory]
    [InlineData(PdfPrintPageSubset.Odd, "5,1,3")]
    [InlineData(PdfPrintPageSubset.Even, "2,4")]
    public void SubsetUsesOriginalPageNumbersAndRetainsSelectedOrder(PdfPrintPageSubset subset, string expected) {
        PdfPrintPlan plan = PdfPrintPlanner.Create(Pages(5), new() {
            InputPath = "snapshot.pdf", Pages = "5,2,1,4,3", PageSubset = subset, PagesPerSheet = 6
        });
        Assert.Equal(expected, string.Join(',', plan.SelectedPages));
        Assert.Equal(plan.SelectedPages, Assert.Single(plan.Sheets).Placements.Select(p => p.PageNumber));
    }

    [Fact]
    public void EmptySubsetAndImpossibleMarginsFailBeforeProducingSheets() {
        PdfDocument document = Pages(1);
        Assert.Throws<ArgumentException>(() => PdfPrintPlanner.Create(document, new() {
            InputPath = "snapshot.pdf", PageSubset = PdfPrintPageSubset.Even
        }));
        Assert.Throws<ArgumentException>(() => PdfPrintPlanner.Create(document, new() {
            InputPath = "snapshot.pdf", PaperSize = new PageSize(100, 120), MarginLeft = 70, MarginRight = 30
        }));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfPrintPlanner.Create(document, new() {
            InputPath = "snapshot.pdf", ScaleMode = PdfPrintScaleMode.Custom, CustomScalePercent = double.NaN
        }));
    }

    [Fact]
    public async Task AsynchronousPlanningSnapshotsNewLayoutSettingsBeforeLoadingTheDocument() {
        var loaded = new TaskCompletionSource<PdfDocument>(TaskCreationOptions.RunContinuationsAsynchronously);
        var request = new PdfPrintPlanRequest {
            InputPath = "snapshot.pdf", PagesPerSheet = 6, PageSubset = PdfPrintPageSubset.Odd,
            ScaleMode = PdfPrintScaleMode.Custom, CustomScalePercent = 20, Alignment = PdfPrintAlignment.BottomRight,
            ColorMode = PdfPrintColorMode.Grayscale, MarginLeft = 24, MarginTop = 30, MarginRight = 36, MarginBottom = 42
        };
        Task<PdfPrintPlan> planning = PdfPrintPlanner.CreateAsync(request, (_, _, _) => loaded.Task);
        request.PagesPerSheet = 1; request.PageSubset = PdfPrintPageSubset.Even;
        request.CustomScalePercent = 100; request.Alignment = PdfPrintAlignment.TopLeft;
        request.ColorMode = PdfPrintColorMode.Color;
        request.MarginLeft = request.MarginTop = request.MarginRight = request.MarginBottom = 0;
        loaded.SetResult(Pages(3));
        PdfPrintPlan plan = await planning;
        Assert.Equal(new[] { 1, 3 }, plan.SelectedPages);
        Assert.Equal(PdfPrintColorMode.Grayscale, plan.ColorMode);
        PdfPrintSheet sheet = Assert.Single(plan.Sheets);
        Assert.All(sheet.Placements, p => {
            Assert.Equal(0.2, p.Scale);
            Assert.Equal(p.SlotX + p.SlotWidth - p.Width, p.X, 6);
            Assert.Equal(p.SlotY + p.SlotHeight - p.Height, p.Y, 6);
        });
        Assert.Equal(24, sheet.Placements[0].SlotX);
        Assert.Equal(30, sheet.Placements[0].SlotY);
    }

    private static PdfDocument Pages(int count) => PdfDocument.Create(c => {
        for (int page = 0; page < count; page++) c.Page(p => p.Size(100, 120).Margin(0));
    });
}
