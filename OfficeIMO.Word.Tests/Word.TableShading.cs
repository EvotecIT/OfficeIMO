using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void TableCellShading_AutomaticFillHasNoRgbColorAndRoundTrips() {
        using WordDocument source = WordDocument.Create();
        WordTableCell cell = source.AddTable(1, 1).Rows[0].Cells[0];
        cell.ShadingFillColorHex = "auto";
        Assert.Equal("AUTO", cell.ShadingFillColorHex);
        Assert.Null(cell.ShadingFillColor);
        cell.ShadingFillColorHex = cell.ShadingFillColorHex;
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Equal("AUTO", loaded.Tables[0].Rows[0].Cells[0].ShadingFillColorHex);
        Assert.Null(loaded.Tables[0].Rows[0].Cells[0].ShadingFillColor);
        Assert.Empty(loaded.DocumentValidationErrors);
    }

    [Theory]
    [InlineData("direct-shading")]
    [InlineData("no-style")]
    [InlineData("no-properties")]
    public void TableShading_ImportedTablesWithoutAStyleExposeSettingsInSchemaOrder(string kind) {
        using WordDocument source = WordDocument.Create();
        WordTable original = source.AddTable(1, 1);
        original.Rows[0].Cells[0].Paragraphs[0].Text = "Imported table";
        original.Rows[0].Cells[0].ShadingFillColorHex = "FFFF00";
        if (kind == "no-properties") original._tableProperties!.Remove();
        else {
            original._tableProperties!.TableStyle?.Remove();
            if (kind == "direct-shading") original._tableProperties.AddChild(
                new W.Shading { Val = W.ShadingPatternValues.Clear, Fill = "FF0000" }, true);
        }
        // Settings access supplies missing table properties; existing style-free properties are already valid.
        if (kind != "no-properties") Assert.Empty(source.DocumentValidationErrors);
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        WordTable table = loaded.Tables[0];
        if (kind != "no-properties") Assert.Empty(loaded.DocumentValidationErrors);
        Assert.Null(table._tableProperties?.TableStyle);
        WordTableStyleDetails details = Assert.IsType<WordTableStyleDetails>(table.StyleDetails);
        Assert.Equal(kind == "direct-shading" ? "FF0000" : string.Empty, details.ShadingFillColorHex);
        details.ShadingFillColorHex = "0000FF";
        details.CellSpacing = 120;
        using WordDocument roundtrip = WordDocument.Load(new MemoryStream(loaded.ToBytes()));
        Assert.Equal("0000FF", roundtrip.Tables[0].StyleDetails!.ShadingFillColorHex);
        Assert.Equal((short)120, roundtrip.Tables[0].StyleDetails!.CellSpacing);
        Assert.Null(roundtrip.Tables[0]._tableProperties!.TableStyle);
        Assert.IsType<W.TableProperties>(roundtrip.Tables[0]._table.ChildElements[0]);
        Assert.Equal("FFFF00", roundtrip.Tables[0].Rows[0].Cells[0].ShadingFillColorHex);
        Assert.Equal("Imported table", roundtrip.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
        Assert.Empty(roundtrip.DocumentValidationErrors);
    }

    [Fact]
    public void TableShading_PublicSettingsRoundTripWithoutChangingCellsOrMargins() {
        using WordDocument source = WordDocument.Create();
        WordTable table = source.AddTable(2, 2, WordTableStyle.TableGrid);
        table.StyleDetails!.CellSpacing = 120;
        table.StyleDetails.MarginDefaultBottomWidth = 80;
        table.Rows[0].Cells[0].ShadingFillColorHex = "FFFFFF";
        table.StyleDetails.ShadingFillColorHex = "#80c0ff";
        table.StyleDetails.ShadingPattern = WordShadingPattern.Percent20;
        table.StyleDetails.ShadingColor = Color.Blue;
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        WordTableStyleDetails style = loaded.Tables[0].StyleDetails!;
        Assert.Equal("80C0FF", style.ShadingFillColorHex);
        Assert.Equal(Color.FromRgb(128, 192, 255), style.ShadingFillColor);
        Assert.Equal(WordShadingPattern.Percent20, style.ShadingPattern);
        Assert.Equal("0000FF", style.ShadingColorHex);
        Assert.Equal(Color.Blue, style.ShadingColor);
        Assert.Equal((short)120, style.CellSpacing);
        Assert.Equal((short)80, style.MarginDefaultBottomWidth);
        Assert.Equal("FFFFFF", loaded.Tables[0].Rows[0].Cells[0].ShadingFillColorHex);
        Assert.Empty(loaded.DocumentValidationErrors);
    }

    [Theory]
    [InlineData("fill")]
    [InlineData("color")]
    public void TableShading_ExplicitRgbReplacesOnlyItsThemeChannel(string channel) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        var shading = new W.Shading { Val = W.ShadingPatternValues.Percent20,
            Fill = "FFFFFF", ThemeFill = W.ThemeColorValues.Accent1, ThemeFillTint = "80", ThemeFillShade = "40",
            Color = "000000", ThemeColor = W.ThemeColorValues.Accent2, ThemeTint = "60", ThemeShade = "20" };
        table._tableProperties!.AddChild(shading, true);
        WordTableStyleDetails style = table.StyleDetails!;
        if (channel == "fill") {
            style.ShadingFillColor = Color.Red;
            Assert.Null(shading.ThemeFill); Assert.Null(shading.ThemeFillTint); Assert.Null(shading.ThemeFillShade);
            Assert.Equal(W.ThemeColorValues.Accent2, shading.ThemeColor!.Value);
            Assert.Equal("000000", style.ShadingColorHex);
        } else {
            style.ShadingColorHex = "#0000ff";
            Assert.Null(shading.ThemeColor); Assert.Null(shading.ThemeTint); Assert.Null(shading.ThemeShade);
            Assert.Equal(W.ThemeColorValues.Accent1, shading.ThemeFill!.Value);
            Assert.Equal("FFFFFF", style.ShadingFillColorHex);
        }
        Assert.Equal(WordShadingPattern.Percent20, style.ShadingPattern);
        Assert.Empty(document.DocumentValidationErrors);
    }

    [Theory]
    [InlineData("fill")]
    [InlineData("pattern")]
    [InlineData("color")]
    public void TableShading_ClearingFollowsTheDocumentedChannelContract(string channel) {
        using WordDocument source = WordDocument.Create();
        WordTableStyleDetails style = source.AddTable(1, 1).StyleDetails!;
        style.ShadingFillColor = Color.Red;
        style.ShadingColor = Color.Blue;
        style.ShadingPattern = WordShadingPattern.Percent20;
        if (channel == "fill") style.ShadingFillColor = null;
        else if (channel == "pattern") style.ShadingPattern = null;
        else style.ShadingColor = null;
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        style = loaded.Tables[0].StyleDetails!;
        Assert.Equal(string.Empty, style.ShadingColorHex);
        if (channel == "color") {
            Assert.Equal(Color.Red, style.ShadingFillColor);
            Assert.Equal(WordShadingPattern.Percent20, style.ShadingPattern);
        } else {
            Assert.Equal(string.Empty, style.ShadingFillColorHex);
            Assert.Null(style.ShadingFillColor); Assert.Null(style.ShadingPattern);
        }
        Assert.Empty(loaded.DocumentValidationErrors);
    }

    [Fact]
    public void TableShading_AutomaticFillClearsInheritedGapFillWithoutInventingRgb() {
        using WordDocument source = WordDocument.Create();
        WordTable table = CreateNativeGapShadingTable(source, "inherited");
        table.StyleDetails!.ShadingFillColorHex = "AUTO";
        table.StyleDetails.ShadingColorHex = "auto";
        using WordDocument loaded = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Equal("AUTO", loaded.Tables[0].StyleDetails!.ShadingFillColorHex);
        Assert.Null(loaded.Tables[0].StyleDetails!.ShadingFillColor);
        Assert.Null(loaded.Tables[0].StyleDetails!.ShadingColor);
        Assert.NotEmpty(loaded.ToBytes(WordFileFormat.Doc, new WordSaveOptions { LossPolicy = OfficeIMO.OfficeConversionLossPolicy.Allow }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableShading_InvalidRgbLeavesExistingSettingsIntact(bool foreground) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        table.StyleDetails!.ShadingFillColor = Color.Red;
        string xml = table._tableProperties!.OuterXml;
        Assert.Throws<ArgumentException>(() => {
            if (foreground) table.StyleDetails.ShadingColorHex = "FF112233";
            else table.StyleDetails.ShadingFillColorHex = "not-a-color";
        });
        Assert.Equal(xml, table._tableProperties.OuterXml);
    }

    [Fact]
    public void TableShading_VisibleNativeDocGapsRemainExplicitlyUnsupported() {
        using WordDocument document = WordDocument.Create();
        WordTableStyleDetails style = document.AddTable(1, 1).StyleDetails!;
        style.CellSpacing = 120;
        style.ShadingFillColorHex = "80C0FF";
        var error = Assert.Throws<NotSupportedException>(() => document.ToBytes(WordFileFormat.Doc,
            new WordSaveOptions { LossPolicy = OfficeIMO.OfficeConversionLossPolicy.Allow }));
        Assert.Contains("table gap shading", error.Message);
    }
}
