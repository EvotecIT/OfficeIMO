using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("WordTableLastRowGreenNative.doc", "00FF00")]
    [InlineData("WordTableFirstRowRedNative.doc", "FF0000")]
    [InlineData("WordTableLastRowNilNative.doc", null)]
    public void LegacyDoc_WordProducedRowBorders_PreserveRowDifferences(string fixture, string? color) {
        using WordDocument document = WordDocument.Load(Path.Combine(AppContext.BaseDirectory, "Documents", fixture));
        var table = document.Tables[0];
        int changedRow = fixture.Contains("FirstRow") ? 0 : 1;
        Assert.Equal(color == null ? WordBorderStyle.Nil : WordBorderStyle.Single,
            table.Rows[changedRow].Cells[0].Borders.LeftStyle);
        Assert.Equal(color, table.Rows[changedRow].Cells[0].Borders.LeftColorHex);
        Assert.Equal(WordBorderStyle.Single, table.Rows[1 - changedRow].Cells[0].Borders.LeftStyle);
        Assert.Equal(color == null ? WordBorderStyle.Nil : WordBorderStyle.Single,
            table.StyleDetails!.GetBorderProperties(changedRow == 0 ? WordTableBorderSide.Top : WordTableBorderSide.Bottom).Style);
        Assert.Equal(color, table.StyleDetails!.GetBorderProperties(
            changedRow == 0 ? WordTableBorderSide.Top : WordTableBorderSide.Bottom).ColorHex);
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx,
            new WordSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow })));
        Assert.Equal(table.Rows[changedRow].Cells[0].Borders.LeftStyle,
            restored.Tables[0].Rows[changedRow].Cells[0].Borders.LeftStyle);
        Assert.Equal(color, restored.Tables[0].Rows[changedRow].Cells[0].Borders.LeftColorHex);
    }

    [Theory]
    [InlineData("WordTableNestedNilNative.doc", WordBorderStyle.Nil, null)]
    [InlineData("WordTableNestedRgbNative.doc", WordBorderStyle.Double, "12AB34")]
    public void LegacyDoc_WordProducedNestedBorders_PreserveExplicitCellOverrides(string fixture, WordBorderStyle style, string? color) {
        using WordDocument document = WordDocument.Load(Path.Combine(AppContext.BaseDirectory, "Documents", fixture));
        var nested = Assert.Single(document.Tables[0].Rows[0].Cells[0].NestedTables);
        Assert.Equal(WordBorderStyle.Single, nested.StyleDetails!.GetBorderProperties(WordTableBorderSide.Top).Style);
        Assert.Equal(style, nested.Rows[0].Cells[0].Borders.TopStyle);
        Assert.Equal(color, nested.Rows[0].Cells[0].Borders.TopColorHex);
        using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx,
            new WordSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow })));
        var cell = Assert.Single(restored.Tables[0].Rows[0].Cells[0].NestedTables).Rows[0].Cells[0];
        Assert.Equal(style, cell.Borders.TopStyle);
        Assert.Equal(color, cell.Borders.TopColorHex);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void LegacyDoc_TableBorderColorOperand_NilSuppressesExistingEdge(int side) {
        var borders = ReadNativeBorderColorOperand(side, new byte[] { 0xFF, 0xFF, 0xFF, 0xFF });
        var values = new[] { borders.Top, borders.Left, borders.Bottom, borders.Right };
        Assert.Equal(LegacyDocTableCellBorderStyle.ExplicitNone, values[side].Style);
        Assert.Equal(LegacyDocTableCellBorderStyle.Single, values[(side + 1) % 4].Style);
    }

    [Fact]
    public void LegacyDoc_TableBorderColorOperand_AutomaticColorRetainsExistingEdge() {
        var borders = ReadNativeBorderColorOperand(0, new byte[] { 0, 0, 0, 0xFF });
        Assert.Equal(LegacyDocTableCellBorderStyle.Single, borders.Top.Style);
        Assert.Null(borders.Top.ColorHex);
    }

    private static LegacyDocTableCellBorders ReadNativeBorderColorOperand(int side, byte[] color) {
        // One-cell TDefTable with four Single TC80 edges, followed by a BrcCvOperand.
        byte[] bytes = new byte[36];
        bytes[0] = 0x08; bytes[1] = 0xD6; bytes[2] = 26; bytes[4] = 1;
        bytes[7] = 0xA0; bytes[8] = 0x05;
        for (int edge = 0; edge < 4; edge++) {
            bytes[13 + edge * 4] = 4;
            bytes[14 + edge * 4] = 1;
            bytes[15 + edge * 4] = 1;
        }
        bytes[29] = (byte)(0x1A + side); bytes[30] = 0xD6; bytes[31] = 4;
        Array.Copy(color, 0, bytes, 32, 4);
        return Assert.Single(LegacyDocParagraphFormattingReader.ReadGrpprl(bytes, 0, bytes.Length)
            .GetTableCellBordersForCellCount(1));
    }
}
