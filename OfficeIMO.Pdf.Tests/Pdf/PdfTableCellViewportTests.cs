using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTableCellViewportTests {
    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void Data_bar_uses_the_unsplit_width_and_is_clipped_at_the_fragment_edge(string mode) {
        var left = new PdfTableCell(string.Empty).WithViewport(new PdfTableCellViewport(200, 24, 100, 24));
        var right = new PdfTableCell(string.Empty).WithViewport(new PdfTableCellViewport(200, 24, 100, 24, offsetX: 100));
        void Configure(PdfTableStyle style) => style.CellDataBars = new Dictionary<(int, int), PdfCellDataBar> {
            [(0, 0)] = new PdfCellDataBar { Color = PdfColor.FromRgb(0, 255, 0), Ratio = 0.5 }
        };
        string first = System.Text.Encoding.ASCII.GetString(Render(mode, left, PdfCellVerticalAlign.Top, Configure));
        string second = System.Text.Encoding.ASCII.GetString(Render(mode, right, PdfCellVerticalAlign.Top, Configure));
        Assert.Contains("24 132 100 24 re f", first);
        Assert.DoesNotContain("0 1 0 rg", second);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void Icon_alignment_can_differ_from_text_and_stays_in_the_original_fragment(string mode) {
        var left = new PdfTableCell("99").WithViewport(new PdfTableCellViewport(200, 24, 100, 24));
        var right = new PdfTableCell("99").WithViewport(new PdfTableCellViewport(200, 24, 100, 24, offsetX: 100));
        var icon = new PdfCellIcon { Color = PdfColor.FromRgb(0, 255, 0), HorizontalAlignment = PdfColumnAlign.Left };
        void Configure(PdfTableStyle style) {
            style.CellAlignments = new Dictionary<(int, int), PdfColumnAlign> { [(0, 0)] = PdfColumnAlign.Right };
            style.CellIcons = new Dictionary<(int, int), PdfCellIcon> { [(0, 0)] = icon };
        }
        byte[] first = Render(mode, left, PdfCellVerticalAlign.Top, Configure);
        byte[] second = Render(mode, right, PdfCellVerticalAlign.Top, Configure);
        Assert.Contains("0 1 0 rg", System.Text.Encoding.ASCII.GetString(first));
        Assert.DoesNotContain("0 1 0 rg", System.Text.Encoding.ASCII.GetString(second));
        using PdfPigDocument firstPdf = PdfPigDocument.Open(first);
        using PdfPigDocument secondPdf = PdfPigDocument.Open(second);
        Assert.DoesNotContain("99", firstPdf.GetPage(1).Text);
        Assert.Contains("99", secondPdf.GetPage(1).Text);
        Assert.Equal(PdfColumnAlign.Left, icon.Clone().HorizontalAlignment);
        Assert.Throws<ArgumentOutOfRangeException>(() => icon.HorizontalAlignment = (PdfColumnAlign)99);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    public void Oversized_viewport_row_rejects_implicit_splitting_that_would_repeat_its_geometry(string mode) {
        var cell = new PdfTableCell("Fragment").WithViewport(new PdfTableCellViewport(100, 48, 100, 24));
        var adjacent = new PdfTableCell(string.Join("\n", Enumerable.Repeat("Long neighbour", 40)));
        ArgumentException error = Assert.Throws<ArgumentException>(() => Render(mode, cell, PdfCellVerticalAlign.Top, style => {
            style.FixedRowHeights = null;
            style.ColumnWidthPoints = new List<double?> { 100, 80 };
            style.RowMinHeights = new List<double?> { 24 };
        }, adjacent));
        Assert.Contains("viewport", error.Message);
    }
    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void Wrapped_middle_aligned_text_keeps_its_line_assignment_across_fragments(string mode) {
        const string text = "First\nSecond\nThird\nFourth";
        var upper = new PdfTableCell(text).WithViewport(new PdfTableCellViewport(100, 100, 100, 50));
        var lower = new PdfTableCell(text).WithViewport(new PdfTableCellViewport(100, 100, 100, 50, offsetY: 50));
        using PdfPigDocument first = PdfPigDocument.Open(Render(mode, upper, PdfCellVerticalAlign.Middle,
            style => style.FixedRowHeights = new List<double?> { 50 }));
        using PdfPigDocument second = PdfPigDocument.Open(Render(mode, lower, PdfCellVerticalAlign.Middle,
            style => style.FixedRowHeights = new List<double?> { 50 }));
        Assert.Contains("First", first.GetPage(1).Text);
        Assert.Contains("Second", first.GetPage(1).Text);
        Assert.DoesNotContain("Third", first.GetPage(1).Text);
        Assert.DoesNotContain("Fourth", first.GetPage(1).Text);
        Assert.Contains("Third", second.GetPage(1).Text);
        Assert.Contains("Fourth", second.GetPage(1).Text);
        Assert.DoesNotContain("First", second.GetPage(1).Text);
        Assert.DoesNotContain("Second", second.GetPage(1).Text);
    }
    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    public void Diagonal_border_keeps_the_full_cell_slope_inside_a_fragment(string mode, bool rounded) {
        var cell = new PdfTableCell(string.Empty).WithViewport(new PdfTableCellViewport(100, 48, 100, 24));
        byte[] bytes = Render(mode, cell, PdfCellVerticalAlign.Top, style => {
            style.CornerRadius = rounded ? 4 : 0;
            style.CellBorders = new Dictionary<(int, int), PdfCellBorder> { [(0, 0)] = new PdfCellBorder {
                Top = false, Right = false, Bottom = false, Left = false, DiagonalDown = true, Width = 1, Color = PdfColor.Black
            } };
        });
        string content = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("24 156 m 124 108 l S", content);
        Assert.Contains("24 132 100 24 re W n", content);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void Viewport_rejects_interactive_cell_fields_that_require_separate_placement() {
        var viewport = new PdfTableCellViewport(100, 48, 100, 24);
        var checkBox = new PdfTableCell("Check", checkBoxes: new[] { new PdfTableCellCheckBox("check") });
        var field = new PdfTableCell("Field", formFields: new[] { PdfTableCellFormField.TextField("field") });
        Assert.Throws<ArgumentException>(() => checkBox.WithViewport(viewport));
        Assert.Throws<ArgumentException>(() => field.WithViewport(viewport));
        Assert.Null(checkBox.WithViewport(null).Viewport);
        Assert.Single(checkBox.WithViewport(null).CheckBoxes);
    }
    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void Vertical_fragments_keep_bottom_alignment_in_the_unsplit_cell(string mode) {
        var top = new PdfTableCell("Bottom").WithViewport(new PdfTableCellViewport(100, 48, 100, 24));
        var bottom = new PdfTableCell("Bottom").WithViewport(new PdfTableCellViewport(100, 48, 100, 24, offsetY: 24));
        using PdfPigDocument first = PdfPigDocument.Open(Render(mode, top, PdfCellVerticalAlign.Bottom));
        using PdfPigDocument second = PdfPigDocument.Open(Render(mode, bottom, PdfCellVerticalAlign.Bottom));
        Assert.DoesNotContain("Bottom", first.GetPage(1).Text);
        Assert.Contains("Bottom", second.GetPage(1).Text);
        Assert.Equal(1, first.NumberOfPages);
        Assert.Equal(1, second.NumberOfPages);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void Horizontal_fragments_keep_left_aligned_text_at_its_original_origin(string mode) {
        var left = new PdfTableCell("Left").WithViewport(new PdfTableCellViewport(200, 24, 100, 24));
        var right = new PdfTableCell("Left").WithViewport(new PdfTableCellViewport(200, 24, 100, 24, offsetX: 100));
        using PdfPigDocument first = PdfPigDocument.Open(Render(mode, left, PdfCellVerticalAlign.Top));
        using PdfPigDocument second = PdfPigDocument.Open(Render(mode, right, PdfCellVerticalAlign.Top));
        Assert.Contains("Left", first.GetPage(1).Text);
        Assert.DoesNotContain("Left", second.GetPage(1).Text);
        Assert.InRange(first.GetPage(1).Letters[0].StartBaseLine.X, 23.9, 24.1);
    }

    [Fact]
    public void Viewport_survives_cell_copy_operations_and_rejects_unusable_geometry() {
        var viewport = new PdfTableCellViewport(200, 48, 100, 24, 100, 24);
        var original = new PdfTableCell("Linked", linkUri: "https://example.com/", namedDestinationName: "Source").WithViewport(viewport);
        PdfTableCell copied = original.WithNoWrap().WithNamedDestination("Fragment");
        Assert.Same(viewport, copied.Viewport);
        Assert.Equal("https://example.com/", copied.LinkUri);
        Assert.Equal("Fragment", copied.NamedDestinationName);
        Assert.Equal("Linked", original.Text);
        Assert.Equal("Source", original.NamedDestinationName);
        Assert.Null(copied.WithViewport(null).Viewport);
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTableCellViewport(100, 24, 100, 24, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTableCellViewport(100, 24, 100, 24, offsetY: 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTableCellViewport(double.PositiveInfinity, 24, 100, 24));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTableCellViewport(1E300, 24, 1E-300, 24));
    }

    private static byte[] Render(string mode, PdfTableCell cell, PdfCellVerticalAlign alignment, Action<PdfTableStyle>? configure = null, PdfTableCell? adjacentCell = null) {
        var options = new PdfOptions { PageWidth = 240, PageHeight = 180,
            MarginTop = 24, MarginBottom = 24, MarginLeft = 24, MarginRight = 24,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 10 };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.FontSize = 10;
        style.LineHeight = 1.2;
        style.ColumnWidthPoints = new List<double?> { 100 };
        style.FixedRowHeights = new List<double?> { 24 };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { alignment };
        style.CellPaddingX = 0;
        style.CellPaddingY = 0;
        style.CellSpacing = 0;
        style.SpacingBefore = 0;
        style.SpacingAfter = 0;
        configure?.Invoke(style);
        var cells = new[] { adjacentCell == null ? new[] { cell } : new[] { cell, adjacentCell } };
        return mode switch {
            "flow" => PdfDocument.Create(options).Table(cells, style: style).ToBytes(),
            "column" => PdfDocument.Create(options).Compose(compose => compose.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.Table(cells, style: style)))))).ToBytes(),
            "canvas" => PdfDocument.Create(options).Canvas(canvas => canvas.Table(cells, 24, 24, 100, 24, style)).ToBytes(),
            _ => throw new ArgumentOutOfRangeException(nameof(mode))
        };
    }
}
