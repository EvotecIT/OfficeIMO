using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LegacyDoc_WordProducedTableGrid_PreservesEditableTableBorderDefaults() {
        // Two synthetic Table Grid tables saved by Microsoft Word 16 in Word 97-2003 format.
        string fixture = Path.Combine(AppContext.BaseDirectory, "Documents", "WordTableGridNative.doc");
        using WordDocument document = WordDocument.Load(fixture);
        Assert.Equal(2, document.Tables.Count);
        WordTable table = document.Tables[0];
        foreach (WordTableBorderSide side in Enum.GetValues(typeof(WordTableBorderSide))) {
            var border = table.StyleDetails!.GetBorderProperties(side);
            Assert.Equal(WordBorderStyle.Single, border.Style);
            Assert.Equal(4U, border.Size);
        }

        table.StyleDetails!.SetCustomBorders(topStyle: WordBorderStyle.Double, topSize: 12,
            topColor: OfficeIMO.Drawing.OfficeColor.Red);
        // The Word-produced fixture also carries preserve-only metadata; this test qualifies its table projection.
        using WordDocument reloaded = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx,
            new WordSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow })));
        var top = reloaded.Tables[0].StyleDetails!.GetBorderProperties(WordTableBorderSide.Top);
        Assert.Equal(WordBorderStyle.Double, top.Style);
        Assert.Equal(12U, top.Size);
        Assert.Equal("FF0000", top.ColorHex);
        Assert.Equal(WordBorderStyle.Single,
            reloaded.Tables[0].StyleDetails!.GetBorderProperties(WordTableBorderSide.InsideHorizontal).Style);
    }

    [Theory]
    [InlineData("WordTableCellColorNative.doc", WordBorderStyle.Double, "12AB34", 12U)]
    [InlineData("WordTableCellNilNative.doc", WordBorderStyle.Nil, null, null)]
    public void LegacyDoc_WordProducedTableCellBorder_OverridesInheritedGrid(
        string name, WordBorderStyle style, string? color, uint? size) {
        using WordDocument document = WordDocument.Load(Path.Combine(AppContext.BaseDirectory, "Documents", name));
        var cell = document.Tables[0].Rows[0].Cells[0];
        Assert.Equal(style, cell.Borders.TopStyle);
        Assert.Equal(color, cell.Borders.TopColorHex);
        Assert.Equal(size, cell.Borders.TopSize);
        Assert.Equal(WordBorderStyle.Single, document.Tables[0].StyleDetails!.GetBorderProperties(WordTableBorderSide.Top).Style);
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx,
            new WordSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow })));
        var border = restored.Tables[0].Rows[0].Cells[0].Borders;
        Assert.Equal(style, border.TopStyle);
        Assert.Equal(color, border.TopColorHex);
        Assert.Equal(size, border.TopSize);
    }

    [Theory]
    [InlineData("WordTableDerivedBordersNative.doc", WordBorderStyle.Nil)]
    [InlineData("WordTableCustomBordersNative.doc", WordBorderStyle.Dotted)]
    public void LegacyDoc_WordProducedTableBorders_PreserveDerivedAndDirectDefinitions(string name, WordBorderStyle bottom) {
        using WordDocument document = WordDocument.Load(Path.Combine(AppContext.BaseDirectory, "Documents", name));
        var details = document.Tables[0].StyleDetails!;
        var top = details.GetBorderProperties(WordTableBorderSide.Top);
        Assert.Equal(WordBorderStyle.Double, top.Style);
        Assert.Equal(12U, top.Size);
        Assert.Equal("12AB34", top.ColorHex);
        Assert.Equal(bottom, details.GetBorderProperties(WordTableBorderSide.Bottom).Style);
        if (name.Contains("Custom")) {
            Assert.Equal(WordBorderStyle.Dashed, details.GetBorderProperties(WordTableBorderSide.Left).Style);
            Assert.Equal("FF0000", details.GetBorderProperties(WordTableBorderSide.InsideHorizontal).ColorHex);
            Assert.Equal("0000FF", details.GetBorderProperties(WordTableBorderSide.InsideVertical).ColorHex);
        }
    }
}
