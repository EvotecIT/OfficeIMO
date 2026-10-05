using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPanelContinuationTests {
    [Fact]
    public void Panel_FinalParagraphSpacingMovesWithItsClosingPadding() {
        var options = new PdfOptions { PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, MaxGeneratedPages = 3 };
        byte[] bytes = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text(string.Join("\n", Enumerable.Range(1, 4).Select(i => $"Marker{i:D3}"))), style: new PdfParagraphStyle {
                LineSpacing = PdfLineSpacing.Exactly(32), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false })
            .Paragraph(p => p.Text("Marker005"), style: new PdfParagraphStyle {
                LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 6, WidowControl = false }),
            new PdfPanelStyle { PaddingX = 0, PaddingY = 4, SpacingBefore = 0, SpacingAfter = 0, RepeatFragmentDecoration = false }).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal(4, pdf.GetPage(1).GetWords().Count(word => word.Text.StartsWith("Marker", StringComparison.Ordinal)));
        Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "Marker005");
    }

    [Theory]
    [InlineData("body")]
    [InlineData("columns")]
    [InlineData("row")]
    public void Panel_UnscaledFinalImageRejectsInsufficientClosingSpace(string layout) {
        var options = new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, MaxGeneratedPages = 3
        };
        var panelStyle = new PdfPanelStyle { PaddingX = 0, PaddingY = 4, SpacingBefore = 0, SpacingAfter = 0, RepeatFragmentDecoration = false };
        void Compose(PdfContentBuilder content) => content.Panel(panel => panel
            .Paragraph(p => p.Text("Marker"), style: new PdfParagraphStyle { LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0 })
            .Image(
            PdfPngTestImages.CreateRgbPng(2, 2), 40, 160, style: new PdfImageStyle { ScaleDownToFit = false }), panelStyle);
        PdfDocument document = PdfDocument.Create(options);
        if (layout == "row") document.Row(row => row.RelativeColumn(Compose));
        else if (layout == "columns") document.Columns(Compose, new PdfMultiColumnOptions { BalanceLastPage = false });
        else document.Panel(panel => panel.Paragraph(p => p.Text("Marker"), style: new PdfParagraphStyle {
            LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0 }).Image(PdfPngTestImages.CreateRgbPng(2, 2), 40, 160,
            style: new PdfImageStyle { ScaleDownToFit = false }), panelStyle);
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData("body", "heading")]
    [InlineData("row", "heading")]
    [InlineData("body", "table")]
    [InlineData("row", "table")]
    [InlineData("body", "list")]
    [InlineData("row", "list")]
    [InlineData("body", "nested")]
    [InlineData("row", "nested")]
    public void Panel_FinalContentKeepsClosingSpaceAcrossSupportedChildren(string layout, string child) {
        var options = new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, MaxGeneratedPages = 3
        };
        var panelStyle = new PdfPanelStyle { PaddingX = 0, PaddingY = 4, SpacingBefore = 0, SpacingAfter = 0, RepeatFragmentDecoration = false };
        string[] markers = Enumerable.Range(1, 6).Select(i => $"Marker{i:D3}").ToArray();
        void Content(PdfContentBuilder content) {
            if (child == "table") content.Table(markers.Select(marker => new[] { marker }), style: new PdfTableStyle {
                HeaderRowCount = 0, FixedRowHeights = Enumerable.Repeat<double?>(26, 6).ToList(),
                CellPaddingX = 0, CellPaddingY = 0, BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 6,
                MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0
            });
            else if (child == "list") content.Bullets(markers, style: new PdfListStyle {
                LineSpacing = PdfLineSpacing.Exactly(26), SpacingBefore = 0, SpacingAfter = 0, ItemSpacing = 0
            });
            else if (child == "nested") content.Panel(inner => inner.Paragraph(p => p.Text(string.Join("\n", markers)),
                style: new PdfParagraphStyle { LineSpacing = PdfLineSpacing.Exactly(25), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false }), panelStyle);
            else {
                content.Paragraph(p => p.Text(string.Join("\n", markers.Take(5))), style: new PdfParagraphStyle {
                    LineSpacing = PdfLineSpacing.Exactly(26), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
                });
                content.Heading(1, markers[5], style: new PdfHeadingStyle {
                    LineSpacing = PdfLineSpacing.Exactly(26), SpacingBefore = 0, SpacingAfter = 0, KeepWithNext = false
                });
            }
        }
        PdfDocument document = PdfDocument.Create(options);
        if (layout == "row") document.Row(row => row.RelativeColumn(column => column.Panel(Content, panelStyle)));
        else document.Panel(Content, panelStyle);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToBytes());
        Assert.True(pdf.NumberOfPages == 2, $"Pages={pdf.NumberOfPages}; first page={pdf.GetPage(1).Text}");
        int firstMarkers = markers.Count(marker => pdf.GetPage(1).Text.Contains(marker, StringComparison.Ordinal));
        Assert.True(firstMarkers == 5, $"Markers={firstMarkers}; first page={pdf.GetPage(1).Text}; second page={pdf.GetPage(2).Text}");
        Assert.Contains("Marker006", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void Columns_ContinuedPanelImageUsesContinuationCapacityDuringBalancing(int depth) {
        var options = new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, MaxGeneratedPages = 4
        };
        var panelStyle = new PdfPanelStyle { PaddingX = 0, PaddingY = 4, FragmentBottomInset = 4, SpacingBefore = 0, SpacingAfter = 0, RepeatFragmentDecoration = false };
        void Content(PdfContentBuilder panel) => panel
            .Paragraph(p => p.Text(string.Join("\n", Enumerable.Range(1, 14).Select(i => $"Marker{i:D3}"))),
                style: new PdfParagraphStyle { LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false })
            .Image(PdfPngTestImages.CreateRgbPng(2, 2), 40, 160, style: new PdfImageStyle { ScaleDownToFit = true });
        byte[] bytes = PdfDocument.Create(options).Columns(columns => columns.Panel(panel => {
            if (depth == 2) panel.Panel(Content, panelStyle);
            else Content(panel);
        }, panelStyle), new PdfMultiColumnOptions { BalanceLastPage = true }).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal(14, pdf.GetPage(1).GetWords().Count(word => word.Text.StartsWith("Marker", StringComparison.Ordinal)));
        Assert.Equal(160 - depth * 4, Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes)).Height, 3);
    }

    [Theory]
    [InlineData(false, 1)]
    [InlineData(true, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 2)]
    public void Row_ContinuedPanelImageUsesItsAvailableCapacity(bool scale, int depth) {
        var options = new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, MaxGeneratedPages = 3
        };
        var panelStyle = new PdfPanelStyle { PaddingX = 0, PaddingY = 4, FragmentBottomInset = 4, SpacingBefore = 0, SpacingAfter = 0, RepeatFragmentDecoration = false };
        void Content(PdfContentBuilder panel) => panel
            .Paragraph(p => p.Text(string.Join("\n", Enumerable.Range(1, 7).Select(i => $"Marker{i:D3}"))),
                style: new PdfParagraphStyle { LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false })
            .Image(PdfPngTestImages.CreateRgbPng(2, 2), 40, scale ? 160 : 160 - depth * 4, style: new PdfImageStyle { ScaleDownToFit = scale });
        byte[] bytes = PdfDocument.Create(options).Row(row => row.RelativeColumn(column => column.Panel(panel => {
            if (depth == 2) panel.Panel(Content, panelStyle);
            else Content(panel);
        }, panelStyle))).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal(160 - depth * 4, Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes)).Height, 3);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("columns")]
    [InlineData("row")]
    public void Panel_ContinuousDecorationRetainsFullClosingPadding(string layout) {
        var options = new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, MaxGeneratedPages = 3
        };
        var style = new PdfPanelStyle { PaddingX = 0, PaddingY = 4, SpacingBefore = 0, SpacingAfter = 0, RepeatFragmentDecoration = false };
        string text = string.Join("\n", Enumerable.Range(1, 6).Select(i => $"Marker{i:D3}"));
        void Compose(PdfContentBuilder content) => content.Panel(panel => panel.Paragraph(p => p.Text(text),
            style: new PdfParagraphStyle { LineSpacing = PdfLineSpacing.Exactly(26), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false }), style);
        PdfDocument document = PdfDocument.Create(options);
        if (layout == "columns") document.Columns(Compose, new PdfMultiColumnOptions { BalanceLastPage = false, Gap = 20 });
        else if (layout == "row") document.Row(row => row.RelativeColumn(Compose));
        else document.Panel(panel => panel.Paragraph(p => p.Text(text), style: new PdfParagraphStyle {
            LineSpacing = PdfLineSpacing.Exactly(26), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
        }), style);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToBytes());
        int page = layout == "columns" ? 1 : 2;
        Assert.Equal(page, pdf.NumberOfPages);
        var marker = Assert.Single(pdf.GetPage(page).GetWords(), word => word.Text == "Marker006");
        Assert.Equal(layout == "columns" ? 160D : 20D, marker.Letters[0].StartBaseLine.X, 3);
        Assert.Equal(5, pdf.GetPage(1).GetWords().Count(word => word.Text.StartsWith("Marker", StringComparison.Ordinal) && word.Letters[0].StartBaseLine.X < 100));
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(6D)]
    public void Panel_BorderOffsetsMovePaintWithoutMovingText(double radius) {
        var options = new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20
        };
        var style = new PdfPanelStyle {
            PaddingX = 0, PaddingY = 4, SpacingBefore = 0, SpacingAfter = 0, CornerRadius = radius,
            LeftBorder = new PdfPanelBorder { Color = PdfColor.Black, Offset = 2 },
            RightBorder = new PdfPanelBorder { Color = PdfColor.Black, Offset = 3 },
            TopBorder = new PdfPanelBorder { Color = PdfColor.Black, Offset = -0.25 },
            BottomBorder = new PdfPanelBorder { Color = PdfColor.Black, Offset = -0.25 }
        };
        byte[] bytes = PdfDocument.Create(options).Panel(panel => panel.Paragraph(p => p.Text("Marker")), style).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = pdf.GetPage(1);
        Assert.Equal(20, Assert.Single(page.GetWords()).BoundingBox.Left, 3);
        var bounds = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(rectangle => rectangle.HasValue).Select(rectangle => rectangle!.Value).ToArray();
        Assert.NotEmpty(bounds);
        Assert.Equal(18, bounds.Min(rectangle => rectangle.Left), 3);
        Assert.Equal(283, bounds.Max(rectangle => rectangle.Right), 3);
        Assert.Equal(179.75, bounds.Max(rectangle => rectangle.Top), 3);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("columns")]
    [InlineData("row")]
    public void Panel_ContinuousDecorationScalesImageWithinItsFragmentInset(string layout) {
        var options = new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, MaxGeneratedPages = 3
        };
        var panelStyle = new PdfPanelStyle {
            PaddingX = 0, PaddingY = 4, FragmentBottomInset = 4,
            SpacingBefore = 0, SpacingAfter = 0, RepeatFragmentDecoration = false
        };
        byte[] image = PdfPngTestImages.CreateRgbPng(2, 2);
        var imageStyle = new PdfImageStyle { ScaleDownToFit = true };
        void Compose(PdfContentBuilder content) => content.Panel(
            panel => panel.Image(image, 40, 160, style: imageStyle), panelStyle);
        PdfDocument document = PdfDocument.Create(options);
        if (layout == "columns") document.Columns(Compose, new PdfMultiColumnOptions { BalanceLastPage = false });
        else if (layout == "row") document.Row(row => row.RelativeColumn(Compose));
        else document.Panel(panel => panel.Image(image, 40, 160, style: imageStyle), panelStyle);

        byte[] bytes = document.ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        var placement = Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(152, placement.Height, 3);
        Assert.Equal(38, placement.Width, 3);
    }

    [Theory]
    [InlineData("body", true, 0D)]
    [InlineData("body", false, 0D)]
    [InlineData("body", false, 4D)]
    [InlineData("columns", true, 0D)]
    [InlineData("columns", false, 0D)]
    [InlineData("columns", false, 4D)]
    [InlineData("row", true, 0D)]
    [InlineData("row", false, 0D)]
    [InlineData("row", false, 4D)]
    public void Panel_ContinuationPreservesItsChosenPaddingPolicy(string layout, bool repeat, double inset) {
        var options = new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFontSize = 12, MaxGeneratedPages = 6
        };
        var panelStyle = new PdfPanelStyle {
            PaddingX = 0, PaddingY = 4, SpacingBefore = 0, SpacingAfter = 0,
            BorderColor = PdfColor.Black, RepeatFragmentDecoration = repeat, FragmentBottomInset = inset
        };
        string text = string.Join("\n", Enumerable.Range(1, 23).Select(index => $"Marker{index:D3}"));
        void Compose(PdfContentBuilder content) => content.Panel(panel => panel.Paragraph(p => p.Text(text),
            style: new PdfParagraphStyle {
                LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
            }), panelStyle);
        PdfDocument document = PdfDocument.Create(options);
        if (layout == "columns") document.Columns(Compose, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = false });
        else if (layout == "row") document.Row(row => row.RelativeColumn(Compose));
        else document.Panel(panel => panel.Paragraph(p => p.Text(text), style: new PdfParagraphStyle {
            LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
        }), panelStyle);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToBytes());
        bool reservesEveryFragment = repeat || inset > 0;
        int expectedPage = layout == "columns" ? (reservesEveryFragment ? 2 : 1) : (reservesEveryFragment ? 3 : 2);
        var marker = Assert.Single(Enumerable.Range(1, pdf.NumberOfPages)
            .SelectMany(page => pdf.GetPage(page).GetWords().Where(word => word.Text == "Marker015")
                .Select(word => (Page: page, Word: word))));
        Assert.Equal(expectedPage, marker.Page);
        var first = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Marker001");
        var continuation = Assert.Single(Enumerable.Range(1, pdf.NumberOfPages)
            .SelectMany(page => pdf.GetPage(page).GetWords()), word => word.Text == "Marker008");
        Assert.Equal(repeat ? 0D : 4D, continuation.BoundingBox.Top - first.BoundingBox.Top, 3);
        var markers = Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords())
            .Where(word => word.Text.StartsWith("Marker", StringComparison.Ordinal)).Select(word => word.Text)
            .OrderBy(value => value, StringComparer.Ordinal);
        Assert.Equal(Enumerable.Range(1, 23).Select(index => $"Marker{index:D3}"), markers);
    }
}
