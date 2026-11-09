using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public sealed class WordAuthoringChainsTests {
    [Fact]
    public void TypographyAndThemeColorPersistThroughTheRunHandle() {
        using var document = WordDocument.Create();
        var run = document.AddParagraph("Styled");
        Assert.Same(run, run.SetBold().SetItalic().SetUnderline(WordUnderlineStyle.Double)
            .SetStrikethrough().SetStrikethrough(false).SetDoubleStrikethrough()
            .SetSuperscript().SetSubscript().SetFontFamily("Aptos")
            .SetFontSize(OfficeLength.Points(13)).SetColor(OfficeColor.Accent3)
            .SetAlignment(WordParagraphAlignment.Center));

        byte[] bytes = document.ToBytes();
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        W.RunProperties native = package.MainDocumentPart!.Document!.Descendants<W.RunProperties>().Single();
        Assert.True(native.Bold != null && native.Italic != null);
        Assert.Null(native.Strike);
        Assert.NotNull(native.DoubleStrike);
        Assert.Equal(W.UnderlineValues.Double, native.Underline!.Val!.Value);
        Assert.Equal(W.VerticalPositionValues.Subscript, native.VerticalTextAlignment!.Val!.Value);
        Assert.Equal(W.ThemeColorValues.Accent3, native.Color!.ThemeColor!.Value);
        using var loaded = WordDocument.Load(new MemoryStream(bytes));
        var actual = loaded.Paragraphs.Single();
        Assert.Equal(OfficeColor.Accent3, actual.Color);
        Assert.True(actual.DoubleStrikethrough);
        Assert.Equal("Aptos", actual.FontFamily);
        Assert.Equal(OfficeLength.Points(13), actual.FontSize);
        Assert.Equal(WordParagraphAlignment.Center, actual.ParagraphAlignment);
    }

    [Theory]
    [InlineData("line")]
    [InlineData("area")]
    [InlineData("bar")]
    public void NumericSequencesAreReadOnceAndReturnAnAttachedSeries(string kind) {
        using var document = WordDocument.Create();
        var chart = document.AddChart();
        bool readCategories = false;
        bool readValues = false;
        IEnumerable<string> Categories() {
            Assert.False(readCategories);
            readCategories = true;
            yield return "First";
            yield return "Second";
        }
        IEnumerable<decimal> Values() {
            Assert.False(readValues);
            readValues = true;
            yield return 1.25m;
            yield return 2.5m;
        }
        chart.AddCategories(Categories());
        WordChartSeries series;
        var previousCulture = System.Globalization.CultureInfo.CurrentCulture;
        try {
            System.Globalization.CultureInfo.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL");
            series = kind == "line" ? chart.AddLine("Values", Values(), OfficeColor.Accent2)
                : kind == "area" ? chart.AddArea("Values", Values(), OfficeColor.Accent2)
                : chart.AddBar("Values", Values(), OfficeColor.Accent2);
        } finally {
            System.Globalization.CultureInfo.CurrentCulture = previousCulture;
        }
        Assert.Same(chart, series.Chart);
        series.Name = "Renamed";
        Assert.Equal(new[] { 1.25, 2.5 }, series.Values);
        using var package = WordprocessingDocument.Open(new MemoryStream(document.ToBytes()), false);
        var part = package.MainDocumentPart!.ChartParts.Single();
        Assert.Contains(part.ChartSpace!.Descendants<C.SeriesText>(), text => text.InnerText == "Renamed");
        Assert.Contains(part.ChartSpace!.Descendants<A.SchemeColor>(), color => color.Val!.InnerText == "accent2");
        Assert.True(readCategories && readValues);
    }

    [Fact]
    public void BulkRowsAndTableConfigurationKeepTheTableAndRowHandles() {
        using var document = WordDocument.Create();
        var table = document.AddTable(1, 2);
        var added = table.AddRows(2, 2);
        added[0].Cells[0].Paragraphs[0].Text = "Added row";
        Assert.Same(table, table.AutoFitToContents().AutoFitToWindow().SetFixedWidth(80)
            .SetWidthPercentage(75).SetColumnWidthsPercentage(40, 60).DistributeColumnsEvenly());
        using var loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        var actual = loaded.Tables.Single();
        Assert.Equal(3, actual.RowCount);
        Assert.All(actual.Rows, row => Assert.Equal(2, row.Cells.Count));
        Assert.Equal("Added row", actual.Rows[1].Cells[0].Paragraphs[0].Text);
    }
}
