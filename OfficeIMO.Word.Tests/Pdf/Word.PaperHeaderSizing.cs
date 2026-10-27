using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPaperHeaderSizingTests {
    [Theory]
    [InlineData("Arial")]
    [InlineData("Helvetica")]
    [InlineData("Courier")]
    public void WideHeadersSerializeWithinThePageAndRenderCompoundWordsIntact(string family) {
        using WordDocument document = WordDocument.Create();
        var style = new WordPaperStyle { FontFamily = family };
        string[] headers = { "Domain", "Authentication", "MTA-STS policy", "TLS-RPT", "Disposition", "Score" };
        WordTable table = document.AddTable(2, headers.Length);
        for (int column = 0; column < headers.Length; column++) {
            style.SetCellText(table.Rows[0].Cells[column], headers[column]);
            style.SetCellText(table.Rows[1].Cells[column], column == 0 ? "example.test" : "1");
        }
        style.ApplyTable(table);
        int available = table.EstimateTableWidthInDxa();
        style.FitTableColumnsToHeaders(table, 45, 5, 5, 5, 25, 15);
        using var serialized = new MemoryStream();
        document.Save(serialized);
        serialized.Position = 0;
        using WordDocument reopened = WordDocument.Load(serialized);
        WordTable saved = reopened.Tables[0];
        Assert.Equal(available, saved.GridColumnWidth.Sum());
        Assert.Equal(available, saved.Width);
        Assert.Equal(WordTableWidthUnit.Dxa, saved.WidthType);
        Assert.Equal(saved.GridColumnWidth, saved.ColumnWidth);
        Assert.All(saved.Rows, row => Assert.Equal(saved.GridColumnWidth, row.Cells.Select(cell => cell.Width!.Value)));
        Assert.True(saved.ColumnWidth[1] > available * .05D);
        Assert.True(saved.ColumnWidth[0] > saved.ColumnWidth[1]);
        byte[] bytes = reopened.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
        string? proof = Environment.GetEnvironmentVariable("OFFICEIMO_PAPER_PROOF_DIR");
        if (!string.IsNullOrWhiteSpace(proof)) {
            Directory.CreateDirectory(proof);
            File.WriteAllBytes(Path.Combine(proof, "wide-headers.pdf"), bytes);
            File.WriteAllBytes(Path.Combine(proof, "wide-headers.docx"), serialized.ToArray());
        }
        PdfReadDocument pdf = PdfReadDocument.Open(bytes);
        Assert.Single(pdf.Pages);
        var spans = pdf.Pages[0].GetTextSpans();
        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var words = independent.GetPage(1).GetWords().ToArray();
        foreach (string word in new[] { "Authentication", "MTA-STS", "TLS-RPT", "Disposition" }) {
            Assert.True(spans.Any(span => span.Text.Contains(word)), word + ": widths=" + string.Join(",", saved.ColumnWidth) + " " + string.Join(" | ", spans.Select(span => $"{span.Text} @ {span.X},{span.Y} size={span.FontSize} width={span.Advance}")));
            PdfTextSpan span = Assert.Single(spans, span => span.Text.Contains(word));
            int column = Array.FindIndex(headers, header => header.Contains(word));
            double left = document.Margins.Left / 20D + saved.ColumnWidth.Take(column).Sum() / 20D;
            Assert.True(span.X >= left, word + " starts outside its column");
            var independentWord = Assert.Single(words, candidate => candidate.Text == word);
            Assert.True(independentWord.BoundingBox.Right <= left + saved.ColumnWidth[column] / 20D + .1D,
                $"{family}: [{span.Text}] size={span.FontSize} font={span.BaseFont}: {word} at {span.X:0.###}..{span.X + span.Advance:0.###} exceeds {left:0.###}..{left + saved.ColumnWidth[column] / 20D:0.###}; widths={string.Join(",", saved.ColumnWidth)}");
        }
    }

    [Fact]
    public void SoftHyphenHeaderMinimumIncludesTheFallbackPaintedHyphen() {
        using WordDocument document = WordDocument.Create();
        var style = new WordPaperStyle { FontFamily = "Unavailable paper font 537" };
        WordTable table = document.AddTable(2, 2);
        style.SetCellText(table.Rows[0].Cells[0], "W\u00ADW");
        style.SetCellText(table.Rows[0].Cells[1], "Notes");
        style.ApplyTable(table, alignNumericColumns: false);
        style.FitTableColumnsToHeaders(table, 0, 1);
        double fallback = OfficeIMO.Drawing.OfficeStandardFontMetrics.MeasureFallbackWidth("W-W", 9, OfficeIMO.Drawing.OfficeFontStyle.Bold);
        Assert.Equal(fallback, OfficeIMO.Drawing.OfficeStandardFontMetrics.MeasureFallbackWidth("W\u00ADW", 9, OfficeIMO.Drawing.OfficeFontStyle.Bold));
        Assert.True(table.ColumnWidth[0] >= fallback * 20 + 2 * style.TableCellHorizontalPadding);
    }

    [Fact]
    public void SizingPreservesANarrowTableAndRelativeWidthsWhenMinimaAlreadyFit() {
        using WordDocument document = WordDocument.Create();
        var style = new WordPaperStyle();
        WordTable table = style.AddKeyValueTable(document, new[] { new System.Collections.Generic.KeyValuePair<string, string>("Host", "example.test") });
        table.Width = 2500;
        int available = table.EstimateTableWidthInDxa();
        style.FitTableColumnsToHeaders(table, 1, 2);
        Assert.Equal(available, table.GridColumnWidth.Sum());
        Assert.InRange(Math.Abs(table.ColumnWidth[1] - table.ColumnWidth[0] * 2), 0, 2);
    }

    [Theory]
    [InlineData("Arial")]
    [InlineData("Unavailable paper font 537")]
    public void ShortHeadersKeepDesiredRatiosAndDoNotChangeBodyTypography(string family) {
        using WordDocument document = WordDocument.Create();
        var style = new WordPaperStyle { FontFamily = family, TableFontSizePoints = 10.5D };
        WordTable table = style.AddKeyValueTable(document,
            new[] { new System.Collections.Generic.KeyValuePair<string, string>("Host", "example.test") }, "Key", "Value");
        WordParagraph body = table.Rows[1].Cells[1].Paragraphs[0];
        body.FontFamily = "Courier New";
        body.FontSizePoints = 12.5D;
        body.ColorHex = "405060";
        style.FitTableColumnsToHeaders(table, 1, 2);
        using var serialized = new MemoryStream();
        document.Save(serialized);
        serialized.Position = 0;
        using WordDocument reopened = WordDocument.Load(serialized);
        WordTable saved = reopened.Tables[0];
        Assert.InRange(Math.Abs(saved.ColumnWidth[1] - saved.ColumnWidth[0] * 2), 0, 2);
        WordParagraph restored = saved.Rows[1].Cells[1].Paragraphs[0];
        Assert.Equal("Courier New", restored.FontFamily);
        Assert.Equal(12.5D, restored.FontSizePoints);
        Assert.Equal("405060", restored.ColorHex);
        Assert.Equal(family, saved.Rows[0].Cells[0].Paragraphs[0].FontFamily);
        Assert.Equal(9D, saved.Rows[0].Cells[0].Paragraphs[0].FontSizePoints);
    }

    [Fact]
    public void ImpossibleHeaderMinimaFailClearlyWithoutChangingTheTable() {
        using WordDocument document = WordDocument.Create();
        var style = new WordPaperStyle();
        WordTable table = document.AddTable(2, 2);
        style.SetCellText(table.Rows[0].Cells[0], new string('W', 120));
        style.SetCellText(table.Rows[0].Cells[1], "TLS-RPT");
        style.ApplyTable(table);
        string before = table._table.OuterXml;
        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(() => style.FitTableColumnsToHeaders(table, 1, 1));
        Assert.Contains("horizontal padding require", failure.Message);
        Assert.Contains("available", failure.Message);
        Assert.Equal(before, table._table.OuterXml);
    }
}
