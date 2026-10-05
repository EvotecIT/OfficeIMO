using System;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfExplicitLineSpacingTests {
    [Theory]
    [InlineData(false, 20D)]
    [InlineData(true, 14D)]
    public void Table_text_shrinking_retains_explicit_paragraph_spacing(bool exact, double expected) {
        var runs = new[] { new PdfTextRun("Q\nThisIdentifierShouldShrinkToFit\nZ", fontSize: 30, font: PdfStandardFont.Courier) };
        var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, lineHeight: 3, fontSize: 8,
            lineSpacing: exact ? PdfLineSpacing.Exactly(14) : PdfLineSpacing.AtLeast(20)) });
        var style = new PdfTableStyle { HeaderRowCount = 0, FontSize = 18, MinimumShrinkFontSize = 7,
            ShrinkTextToFit = true, ColumnWidthPoints = new() { 180 }, CellPaddingX = 0, CellPaddingY = 0 };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(new[] { new[] { cell } }, style: style).ToBytes());
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "Q");
        Assert.InRange(first.FontSize, 7D, 29.99D);
        Assert.Single(letters, letter => letter.Value == "Z");
        var baselines = letters.Select(letter => letter.StartBaseLine.Y).Distinct().OrderByDescending(y => y).ToArray();
        Assert.True(baselines.Length >= 3);
        Assert.All(Enumerable.Range(1, baselines.Length - 1), index => Assert.Equal(expected, baselines[index - 1] - baselines[index], 3));
    }

    [Theory]
    [InlineData(false, 52D)]
    [InlineData(true, 40D)]
    public void Styled_blank_lines_expand_minimum_spacing_and_preserve_exact_spacing(bool exact, double expected) {
        var style = new PdfParagraphStyle { FontSize = 8, SpacingAfter = 0,
            LineSpacing = exact ? PdfLineSpacing.Exactly(20) : PdfLineSpacing.AtLeast(20, 1) };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Paragraph(p =>
            p.FontSize(8).Text("A").LineBreak().FontSize(32).LineBreak().FontSize(8).Text("B"), style: style).ToBytes());
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(expected, Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.Y -
            Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData("flow", "multiple", 10D)]
    [InlineData("column", "multiple", 10D)]
    [InlineData("table", "multiple", 10D)]
    [InlineData("flow", "exact", 14D)]
    [InlineData("column", "exact", 14D)]
    [InlineData("table", "exact", 14D)]
    [InlineData("flow", "minimum", 20D)]
    [InlineData("column", "minimum", 20D)]
    [InlineData("table", "minimum", 20D)]
    [InlineData("flow", "minimumzero", 10D)]
    public void Paragraph_spacing_uses_its_font_size_and_explicit_units_in_every_frame(string frame, string rule, double expected) {
        var style = new PdfParagraphStyle { FontSize = 8, LineHeight = 3, SpacingAfter = 0,
            LineSpacing = rule switch { "multiple" => PdfLineSpacing.Multiple(1.25),
                "exact" => PdfLineSpacing.Exactly(14), "minimumzero" => PdfLineSpacing.AtLeast(0, 1.25),
                _ => PdfLineSpacing.AtLeast(20, 1.25) } };
        using var pdf = PdfPigDocument.Open(Render(frame, style, false));
        var letters = pdf.GetPage(1).Letters;
        double first = letters.First(letter => letter.Value == "A").StartBaseLine.Y;
        double second = letters.First(letter => letter.Value == "B").StartBaseLine.Y;
        double third = letters.First(letter => letter.Value == "C").StartBaseLine.Y;
        Assert.InRange(first - second, expected - 0.01, expected + 0.01);
        Assert.InRange(second - third, expected - 0.01, expected + 0.01);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("table")]
    public void Exact_spacing_stays_fixed_with_larger_runs(string frame) {
        var style = new PdfParagraphStyle { FontSize = 8, SpacingAfter = 0, LineSpacing = PdfLineSpacing.Exactly(14) };
        using var pdf = PdfPigDocument.Open(Render(frame, style, true));
        var letters = pdf.GetPage(1).Letters;
        double first = letters.First(letter => letter.Value == "A").StartBaseLine.Y;
        double second = letters.First(letter => letter.Value == "B").StartBaseLine.Y;
        double third = letters.First(letter => letter.Value == "C").StartBaseLine.Y;
        Assert.InRange(first - second, 13.99, 14.01);
        Assert.InRange(second - third, 13.99, 14.01);
    }

    [Fact]
    public void Exact_spacing_measurement_keeps_a_fitting_paragraph_together() {
        var options = Options(); options.PageHeight = 130; options.MarginTop = 20; options.MarginBottom = 20;
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options).Paragraph(p => p.Text("A\nB\nC\nD\nE\nF"),
            style: new PdfParagraphStyle { FontSize = 8, KeepTogether = true, SpacingAfter = 0, LineSpacing = PdfLineSpacing.Exactly(14) }).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("F", pdf.GetPage(1).Text);
    }

    [Fact]
    public void Paragraph_font_and_spacing_are_snapshotted_for_later_rendering() {
        var style = new PdfParagraphStyle { FontSize = 8, SpacingAfter = 0, LineSpacing = PdfLineSpacing.Exactly(14) };
        var document = PdfDocument.Create(Options()).Paragraph(p => p.Text("A\nB"), style: style);
        style.FontSize = 32; style.LineSpacing = PdfLineSpacing.AtLeast(40);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var letters = pdf.GetPage(1).Letters;
        Assert.InRange(letters.First(letter => letter.Value == "A").StartBaseLine.Y -
            letters.First(letter => letter.Value == "B").StartBaseLine.Y, 13.99, 14.01);
    }

    private static PdfOptions Options() => new() { PageWidth = 300, PageHeight = 300,
        MarginTop = 30, MarginBottom = 30, MarginLeft = 30, MarginRight = 30, DefaultFontSize = 12 };

    private static byte[] Render(string frame, PdfParagraphStyle style, bool mixed) {
        Action<PdfParagraphBuilder> text = p => p.Text("A\n").FontSize(mixed ? 32 : 8).Text("B\n").FontSize(8).Text("C");
        if (frame == "flow") return PdfDocument.Create(Options()).Paragraph(text, style: style).ToBytes();
        if (frame == "column") return PdfDocument.Create(Options()).Compose(c => c.Page(page => page.Content(content =>
            content.Row(row => row.PercentColumn(100, column => column.Paragraph(text, style: style)))))).ToBytes();
        var runs = new[] { new PdfTextRun("A\n", fontSize: 8), new PdfTextRun("B\n", fontSize: mixed ? 32 : 8), new PdfTextRun("C", fontSize: 8) };
        var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs,
            fontSize: style.FontSize, lineSpacing: style.LineSpacing) });
        var tableStyle = TableStyles.Minimal(); tableStyle.HeaderRowCount = 0;
        tableStyle.CellPaddingX = 0; tableStyle.CellPaddingY = 0;
        return PdfDocument.Create(Options()).Table(new[] { new[] { cell } }, style: tableStyle).ToBytes();
    }
}
