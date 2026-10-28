using OfficeIMO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LegacyDoc_ThemedTableReportsPaletteApproximationBeforePublication() {
        using WordDocument source = WordDocument.Create();
        WordTable table = source.AddTable(3, 2, WordTableStyle.OfficeIMOModern);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Value";
        table.Rows[1].Cells[0].Paragraphs[0].Text = "North";
        table.Rows[1].Cells[1].Paragraphs[0].Text = "42";
        string styleBefore = GetStyleDefinition(source, "OfficeIMOModern").OuterXml;

        using var destination = new MemoryStream();
        destination.WriteByte(123);
        OfficeConversionException failure = Assert.Throws<OfficeConversionException>(() =>
            source.Save(destination, new WordSaveOptions { Format = WordFileFormat.Doc }));
        Assert.Equal(new byte[] { 123 }, destination.ToArray());
        Assert.Contains(failure.Report.FidelityDiagnostics,
            diagnostic => diagnostic.Code == "word-doc-color-approximated" &&
                diagnostic.LossKind == OfficeConversionLossKind.Approximation);

        byte[] bytes = source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc, AllowLoss = true });
        using WordDocument restored = WordDocument.Load(new MemoryStream(bytes));
        WordTable native = Assert.Single(restored.Tables);
        Assert.Equal(3, native.Rows.Count);
        Assert.Equal("Name", native.Rows[0].Cells[0].Paragraphs[0].Text);
        Assert.Equal("42", native.Rows[1].Cells[1].Paragraphs[0].Text);
        Assert.Equal(styleBefore, GetStyleDefinition(source, "OfficeIMOModern").OuterXml);
    }

    [Fact]
    public void LegacyDoc_ExactPaletteTableSavesWithoutAcceptingLoss() {
        using WordDocument source = WordDocument.Create();
        source.AddTable(1, 1, WordTableStyle.TableGrid).Rows[0].Cells[0].Paragraphs[0].Text = "Exact";
        byte[] bytes = source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc });
        using WordDocument restored = WordDocument.Load(new MemoryStream(bytes));
        Assert.Equal("Exact", Assert.Single(restored.Tables).Rows[0].Cells[0].Paragraphs[0].Text);
    }
    [Theory]
    [InlineData("paragraph shading")]
    [InlineData("paragraph border")]
    [InlineData("table cell shading")]
    [InlineData("table cell border")]
    [InlineData("page border")]
    public void LegacyDoc_PaletteLossIsExplicitAcrossNativePaint(string kind) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Paint");
        WordTableCell cell = source.AddTable(1, 1, WordTableStyle.TableGrid).Rows[0].Cells[0];
        cell.Paragraphs[0].Text = "Cell";
        switch (kind) {
            case "paragraph shading": paragraph.ShadingFillColorHex = "336699"; break;
            case "paragraph border": paragraph.Borders.TopStyle = WordBorderStyle.Single; paragraph.Borders.TopColorHex = "336699"; break;
            case "table cell shading": cell.ShadingFillColorHex = "336699"; break;
            case "table cell border": cell.Borders.TopStyle = WordBorderStyle.Single; cell.Borders.TopColorHex = "336699"; break;
            case "page border": source.Sections[0].Borders.TopStyle = WordBorderStyle.Single; source.Sections[0].Borders.TopColorHex = "336699"; break;
        }
        OfficeConversionException failure = Assert.Throws<OfficeConversionException>(() =>
            source.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc }));
        var finding = Assert.Single(failure.Report.FidelityDiagnostics);
        Assert.Equal("word-doc-color-approximated", finding.Code);
        Assert.Equal(OfficeConversionLossKind.Approximation, finding.LossKind);
        Assert.Contains("#336699", finding.Message);
        Assert.Contains("#008080", finding.Message);
        using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(
            new WordSaveOptions { Format = WordFileFormat.Doc, AllowLoss = true })));
        WordParagraph nativeParagraph = restored.Paragraphs.First(item => item.Text == "Paint");
        WordTableCell nativeCell = Assert.Single(restored.Tables).Rows[0].Cells[0];
        string? actual = kind switch {
            "paragraph shading" => nativeParagraph.ShadingFillColorHex,
            "paragraph border" => nativeParagraph.Borders.TopColorHex,
            "table cell shading" => nativeCell.ShadingFillColorHex,
            "table cell border" => nativeCell.Borders.TopColorHex,
            "page border" => restored.Sections[0].Borders.TopColorHex,
            _ => null
        };
        Assert.Equal("008080", actual);
    }

    [Theory]
    [InlineData(OfficeCompatibilityMode.StrictNative, false)]
    [InlineData(OfficeCompatibilityMode.PreferEditable, true)]
    public void LegacyDoc_ConversionRetainsTheNativePaletteReport(OfficeCompatibilityMode mode, bool expectedOutput) {
        string stem = Path.Combine(_directoryWithFiles, "PaletteConversion-" + Guid.NewGuid().ToString("N"));
        string input = stem + ".docx", output = stem + ".doc";
        try {
            using (WordDocument source = WordDocument.Create(input)) {
                source.AddTable(2, 1, WordTableStyle.OfficeIMOModern).Rows[0].Cells[0].Paragraphs[0].Text = "Editable";
                source.Save();
            }
            var options = new WordDocumentConversionOptions { CompatibilityMode = mode };
            var result = expectedOutput ? WordDocument.Convert(input, output, options)
                : Assert.Throws<WordDocumentConversionException>(() => WordDocument.Convert(input, output, options)).Result;
            Assert.Equal(expectedOutput, result.Succeeded);
            Assert.Equal(expectedOutput, File.Exists(output));
            Assert.True(result.Report.HasLoss);
            Assert.Contains(result.Report.Diagnostics, finding => finding.Code == "word-doc-color-approximated" &&
                finding.CompatibilityState == (expectedOutput ? OfficeCompatibilityState.Approximated : OfficeCompatibilityState.Blocked));
            Assert.DoesNotContain(result.Report.Diagnostics, finding => finding.Code == "word-legacy-writer-visual-fallback");
            if (expectedOutput) {
                using WordDocument restored = WordDocument.Load(output);
                Assert.Equal("Editable", Assert.Single(restored.Tables).Rows[0].Cells[0].Paragraphs[0].Text);
            }
        } finally {
            DeleteIfExists(input);
            DeleteIfExists(output);
        }
    }

}
