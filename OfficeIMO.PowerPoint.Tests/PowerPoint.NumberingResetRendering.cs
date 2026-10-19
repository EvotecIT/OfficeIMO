using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointNumberingResetRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Saved_PPTX_parent_items_reset_deeper_rendered_counters(bool tableCell) {
        using var saved = new MemoryStream();
        using (var presentation = PowerPointPresentation.Create()) {
            presentation.SlideSize.SetSizePoints(600, 500);
            AddNestedItems(presentation.AddSlide(), tableCell);
            presentation.Save(saved);
        }
        saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        Assert.Empty(reopened.ValidateDocument());
        var snapshot = reopened.Slides[0].CreateVisualSnapshot();
        string[] lines = snapshot.TextSpans.GroupBy(span => Math.Round(span.Y, 2))
            .OrderByDescending(line => line.Key)
            .Select(line => string.Concat(line.OrderBy(span => span.X).Select(span => span.Text)))
            .ToArray();
        Assert.Equal(ExpectedItems(), lines);
    }

    [Fact]
    public void Saved_PPTX_table_number_markers_reach_PDF_text() {
        using var saved = new MemoryStream();
        using (var presentation = PowerPointPresentation.Create()) {
            presentation.SlideSize.SetSizePoints(600, 500);
            AddNestedItems(presentation.AddSlide(), tableCell: true);
            presentation.Save(saved);
        }
        saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        var result = reopened.ToPdfDocumentResult();
        var pdf = PdfCore.PdfReadDocument.Open(result.Value.ToBytes());
        string text = pdf.ExtractText();
        AssertNumberedItemsInOrder(text);
        Assert.Equal(OfficeConversionLossKind.Approximation,
            Assert.Single(result.Report.Warnings, warning => warning.Code == "table-list-layout-approximation").LossKind);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PDF_table_markers_retain_separate_font_and_point_size(bool characterBullet) {
        using var presentation = PowerPointPresentation.Create();
        presentation.SlideSize.SetSizePoints(300, 200);
        var cell = presentation.AddSlide().AddTablePoints(1, 1, 20, 20, 260, 150).GetCell(0, 0);
        cell.FontSize = 12;
        cell.FontName = "Arial";
        var paragraph = cell.Paragraphs[0];
        paragraph.Text = "Separate marker";
        paragraph.Runs[0].FontName = "Arial";
        paragraph.Runs[0].FontSize = 12;
        if (characterBullet) paragraph.SetBullet('*');
        else paragraph.SetNumbered(PowerPointNumberingScheme.ArabicPeriod, 7);
        paragraph.SetBulletFont("Times New Roman").SetBulletSizePoints(11);
        using var saved = new MemoryStream(); presentation.Save(saved); saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        var result = reopened.ToPdfDocumentResult(new PowerPointToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        var spans = PdfCore.PdfReadDocument.Open(result.ToBytes()).Pages[0].GetTextSpans();
        var marker = Assert.Single(spans, span => span.Text.TrimEnd() == (characterBullet ? "*" : "7."));
        var body = Assert.Single(spans, span => span.Text == "Separate marker");
        Assert.Contains("Times", marker.BaseFont ?? string.Empty, StringComparison.Ordinal);
        Assert.Equal(11D, marker.FontSize);
        Assert.Contains("Helvetica", body.BaseFont ?? string.Empty, StringComparison.Ordinal);
        Assert.Equal(12D, body.FontSize);
        Assert.NotEqual(marker.FontResource, body.FontResource);
        Assert.True(marker.X < body.X);
    }

    [Theory]
    [InlineData(false, false, 10)]
    [InlineData(true, false, 12)]
    [InlineData(false, true, 14)]
    public void PDF_table_marker_percent_uses_effective_row_font_default(bool header, bool footer, int expectedSize) {
        using var presentation = PowerPointPresentation.Create();
        var table = presentation.AddSlide().AddTablePoints(1, 1, 20, 20, 260, 150);
        table.FirstRow = header;
        table.LastRow = footer;
        var paragraph = table.GetCell(0, 0).Paragraphs[0];
        paragraph.Text = "Inherited body size";
        paragraph.SetNumbered(PowerPointNumberingScheme.ArabicPeriod, 7);
        paragraph.SetBulletSizePercent(50);
        var options = new PowerPointToPdfOptions { PdfOptions = new PdfCore.PdfOptions {
            DefaultTableStyle = new PdfCore.PdfTableStyle { FontSize = 20, HeaderFontSize = 24, FooterFontSize = 28 }
        }};
        byte[] bytes = presentation.ToPdfBytes(options);
        var spans = PdfCore.PdfReadDocument.Open(bytes).Pages[0].GetTextSpans();
        var marker = Assert.Single(spans, span => span.Text.TrimEnd() == "7.");
        Assert.Equal((double)expectedSize, marker.FontSize);
        var body = Assert.Single(spans, span => span.Text == "Inherited body size");
        Assert.Equal(expectedSize * 2D, body.FontSize);
    }

    private static void AssertNumberedItemsInOrder(string text) {
        int previousEnd = 0;
        foreach (string item in ExpectedItems()) {
            int position = text.IndexOf(item, previousEnd, StringComparison.Ordinal);
            Assert.True(position >= previousEnd, "Missing or out-of-order numbered PDF item: " + item + "\n" + text);
            previousEnd = position + item.Length;
        }
    }

    private static string[] ExpectedItems() => new[] { "7. Item 0", "1. Item 1", "1. Item 2",
        "2. Item 3", "1. Item 4", "8. Item 5", "5. Item 6", "6. Item 7", "9. Item 8", "1. Item 9" };

    private static void AddNestedItems(PowerPointSlide slide, bool tableCell) {
        var box = tableCell ? null : slide.AddTextBoxPoints(string.Empty, 20, 20, 560, 450);
        var cell = tableCell ? slide.AddTablePoints(1, 1, 20, 20, 560, 450).GetCell(0, 0) : null;
        if (box != null) box.FontSize = 12;
        if (cell != null) cell.FontSize = 12;
        int[] levels = { 0, 1, 2, 1, 2, 0, 1, 1, 0, 1 };
        int?[] starts = { 7, 1, 1, null, null, null, 5, null, null, null };
        for (int index = 0; index < levels.Length; index++) {
            var paragraph = index == 0 ? box?.Paragraphs[0] ?? cell!.Paragraphs[0]
                : box?.AddParagraph() ?? cell!.AddParagraph();
            paragraph.Text = "Item " + index;
            paragraph.Level = levels[index];
            if (starts[index] is int start) paragraph.SetNumbered(PowerPointNumberingScheme.ArabicPeriod, start);
            else paragraph.SetNumbered(PowerPointNumberingScheme.ArabicPeriod);
        }
    }
}
