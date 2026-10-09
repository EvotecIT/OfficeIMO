using OfficeIMO.Word.LegacyDoc.Model;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LegacyDoc_DataTableProperties_ApplyModernWidthsAndSkipOuterFallback() {
        byte[] modern = {
            0x16, 0x24, 1, 0x17, 0x24, 1, 0x49, 0x66, 1, 0, 0, 0,
            0x21, 0x76, 0, 2, 0x68, 1, 0x23, 0x76, 0, 1, 0xD0, 7,
            0x23, 0x76, 1, 2, 0xB8, 0x0B, 0x01, 0x96, 0x2C, 1,
            0x02, 0x96, 0x6C, 0, 0x07, 0x94, 0x20, 3
        };
        byte[] outer = { 0x6B, 0x64, 0, 0, 0, 0, 0x07, 0x94, 0x10, 0x27 };
        byte[] resolved = LegacyDocParagraphFormattingReader.ResolveDataProperties(outer, 0, outer.Length, DataRecord(modern));
        LegacyDocParagraphFormat format = LegacyDocParagraphFormattingReader.ReadGrpprl(resolved, 0, resolved.Length, requireComplete: true);
        Assert.Equal(new[] { 2000, 3000 }, format.TableCellWidthsTwips);
        Assert.Equal(192, format.TableLeftIndentTwips);
        Assert.Equal(800, format.TableRowHeightTwips);
        Assert.True(format.IsTableTerminatingParagraph);
    }

    [Fact]
    public void LegacyDoc_DataTableProperties_PreservePrefixAndStopEachContainingArray() {
        byte[] child = { 0x07, 0x94, 0x20, 3, 0x16, 0x24, 1, 0x17, 0x24, 1 };
        byte[] parent = { 0x05, 0x24, 1, 0x6B, 0x64, 16, 0, 0, 0, 0x07, 0x94, 0x10, 0x27 };
        byte[] data = DataRecord(parent).Concat(new byte[] { 0 }).Concat(DataRecord(child)).ToArray();
        byte[] root = { 0x6B, 0x64, 0, 0, 0, 0, 0x05, 0x24, 0 };
        byte[] resolved = LegacyDocParagraphFormattingReader.ResolveDataProperties(root, 0, root.Length, data);
        LegacyDocParagraphFormat format = LegacyDocParagraphFormattingReader.ReadGrpprl(resolved, 0, resolved.Length, requireComplete: true);
        Assert.True(format.KeepLinesTogether);
        Assert.Equal(800, format.TableRowHeightTwips);
    }

    [Fact]
    public void LegacyDoc_DataTableProperties_RejectCyclicAndTruncatedReferencedRecords() {
        byte[] root = { 0x6B, 0x64, 0, 0, 0, 0 };
        byte[] cyclic = DataRecord(new byte[] { 0x6B, 0x64, 0, 0, 0, 0, 0x07, 0x94, 0x20, 3 });
        Assert.Throws<InvalidDataException>(() => LegacyDocParagraphFormattingReader.ResolveDataProperties(root, 0, root.Length, cyclic));
        foreach (byte[] invalid in new[] { Array.Empty<byte>(), new byte[] { 0xFF, 0xFF }, new byte[] { 10, 0, 1 } })
            Assert.Throws<InvalidDataException>(() => LegacyDocParagraphFormattingReader.ResolveDataProperties(root, 0, root.Length, invalid));
    }

    [Fact]
    public void LegacyDoc_HugeParagraphProperties_RequireOnlyPapxPropertyAndStyleZero() {
        byte[] child = { 0x07, 0x94, 0x20, 3, 0x16, 0x24, 1, 0x17, 0x24, 1 };
        byte[] root = { 0x46, 0x66, 0, 0, 0, 0 };
        byte[] data = DataRecord(child);
        Assert.Equal(child, LegacyDocParagraphFormattingReader.ResolveDataProperties(root, 0, root.Length, data));
        Assert.Throws<InvalidDataException>(() => LegacyDocParagraphFormattingReader.ResolveDataProperties(root, 0, root.Length, data, styleIndex: 1));
        byte[] trailing = root.Concat(new byte[] { 0x05, 0x24, 1 }).ToArray();
        Assert.Throws<InvalidDataException>(() => LegacyDocParagraphFormattingReader.ResolveDataProperties(trailing, 0, trailing.Length, data));
        byte[] ignored = new byte[] { 0x05, 0x24, 1 }.Concat(root).Concat(new byte[] { 0x07, 0x94, 0x20, 3 }).ToArray();
        byte[] resolved = LegacyDocParagraphFormattingReader.ResolveDataProperties(ignored, 0, ignored.Length, data);
        Assert.Equal(new byte[] { 0x05, 0x24, 1, 0x07, 0x94, 0x20, 3 }, resolved);
    }

    [Fact]
    public void LegacyDoc_ModernCellMutations_PreserveExistingCellFormattingWhenIndicesShift() {
        byte[] initial = {
            0x21, 0x76, 0, 2, 0xD0, 7,
            0x32, 0xD6, 6, 1, 2, 0x01, 3, 40, 0,
            0x21, 0x76, 1, 1, 0xE8, 3,
            0x22, 0x56, 0, 1
        };
        LegacyDocParagraphFormat format = LegacyDocParagraphFormattingReader.ReadGrpprl(initial, 0, initial.Length, requireComplete: true);
        Assert.Equal(new[] { 1000, 2000 }, format.TableCellWidthsTwips);
        var margins = format.GetTableCellMarginsForCellCount(2);
        Assert.Null(margins[0].TopTwips);
        Assert.Equal(40, margins[1].TopTwips);
    }

    [Theory]
    [InlineData(0, 0, 0)]
    [InlineData(1, 1, 1000)]
    [InlineData(0, 64, 1)]
    [InlineData(0, 2, 20000)]
    public void LegacyDoc_ModernCellInsertion_RejectsInvalidCellRangesOrWidth(int first, int count, int width) {
        byte[] properties = { 0x21, 0x76, (byte)first, (byte)count, (byte)width, (byte)(width >> 8) };
        Assert.Throws<InvalidDataException>(() => LegacyDocParagraphFormattingReader.ReadGrpprl(properties, 0, properties.Length, requireComplete: true));
    }

    private static byte[] DataRecord(byte[] properties) =>
        new[] { (byte)properties.Length, (byte)(properties.Length >> 8) }.Concat(properties).ToArray();

    [Fact]
    public void LegacyDoc_ModernCellBorders_ApplyRangesInOrderAndPreserveOtherEdges() {
        byte[] properties = {
            0x21, 0x76, 0, 2, 0xD0, 7,
            0x2F, 0xD6, 11, 0, 2, 0x01, 0x12, 0xAB, 0x34, 0, 12, 3, 0, 0,
            0x2F, 0xD6, 11, 1, 2, 0x02, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF
        };
        LegacyDocParagraphFormat format = LegacyDocParagraphFormattingReader.ReadGrpprl(properties, 0, properties.Length, requireComplete: true);
        Assert.Equal(2, format.TableCellBorders.Count);
        Assert.All(format.TableCellBorders, cell => {
            Assert.Equal(LegacyDocTableCellBorderStyle.Double, cell.Top.Style);
            Assert.Equal("12AB34", cell.Top.ColorHex);
            Assert.Equal(12, cell.Top.SizeEighthPoints);
            Assert.Equal(LegacyDocTableCellBorderStyle.None, cell.Bottom.Style);
        });
        Assert.Equal(LegacyDocTableCellBorderStyle.None, format.TableCellBorders[0].Left.Style);
        Assert.Equal(LegacyDocTableCellBorderStyle.ExplicitNone, format.TableCellBorders[1].Left.Style);
    }

    [Theory]
    [InlineData(0x5622)]
    [InlineData(0x7623)]
    [InlineData(0xD62F)]
    public void LegacyDoc_ReviewedEmptyCellRanges_PreserveRowProperties(int code) {
        byte[] operation = code == 0x5622 ? new byte[] { 0x22, 0x56, 0, 0 }
            : code == 0x7623 ? new byte[] { 0x23, 0x76, 0, 0, 0xE8, 3 }
            : new byte[] { 0x2F, 0xD6, 11, 0, 0, 1, 0x12, 0xAB, 0x34, 0, 12, 3, 0, 0 };
        byte[] properties = new byte[] { 0x16, 0x24, 1, 0x17, 0x24, 1, 0x21, 0x76, 0, 1, 0xD0, 7 }
            .Concat(operation).Concat(new byte[] { 0x07, 0x94, 0x20, 3 }).ToArray();
        LegacyDocParagraphFormat format = LegacyDocParagraphFormattingReader.ReadGrpprl(properties, 0, properties.Length, requireComplete: true);
        Assert.Equal(new[] { 2000 }, format.TableCellWidthsTwips);
        Assert.True(format.IsTableTerminatingParagraph);
        Assert.Equal(800, format.TableRowHeightTwips);
        Assert.Empty(format.TableCellBorders);
    }

    [Fact]
    public void LegacyDoc_ReviewedZeroWidthColumns_RetainIndicesForFollowingProperties() {
        byte[] properties = DataTableDefinition(new short[] { 0, 0, 2000 }, new byte[40])
            .Concat(new byte[] { 0x23, 0x76, 1, 2, 0xE8, 3, 0x32, 0xD6, 6, 1, 2, 1, 3, 40, 0 }).ToArray();
        LegacyDocParagraphFormat format = LegacyDocParagraphFormattingReader.ReadGrpprl(properties, 0, properties.Length, requireComplete: true);
        Assert.Equal(new[] { 0, 1000 }, format.TableCellWidthsTwips);
        var margins = format.GetTableCellMarginsForCellCount(2);
        Assert.Null(margins[0].TopTwips);
        Assert.Equal(40, margins[1].TopTwips);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void LegacyDoc_ReviewedMissingCellDefinitions_UseDefaults(int definitions) {
        byte[] cellDefinitions = new byte[definitions * 20];
        if (definitions > 0) cellDefinitions[0] = 0x80;
        byte[] properties = DataTableDefinition(new short[] { 0, 2000, 4000 }, cellDefinitions);
        LegacyDocParagraphFormat format = LegacyDocParagraphFormattingReader.ReadGrpprl(properties, 0, properties.Length, requireComplete: true);
        Assert.Equal(new[] { 2000, 2000 }, format.TableCellWidthsTwips);
        Assert.Empty(format.TableCellBorders);
        if (definitions == 0) Assert.Empty(format.TableCellVerticalAlignments);
        else Assert.Equal(new[] { LegacyDocTableCellVerticalAlignment.Center, LegacyDocTableCellVerticalAlignment.Top }, format.TableCellVerticalAlignments);
    }

    [Fact]
    public void LegacyDoc_ReviewedPartialCellDefinition_IsRejected() {
        byte[] properties = DataTableDefinition(new short[] { 0, 2000, 4000 }, new byte[19]);
        Assert.Throws<InvalidDataException>(() => LegacyDocParagraphFormattingReader.ReadGrpprl(properties, 0, properties.Length, requireComplete: true));
    }

    private static byte[] DataTableDefinition(short[] edges, byte[] cells) {
        int count = edges.Length - 1;
        int length = 5 + edges.Length * 2 + cells.Length;
        var result = new byte[length];
        result[0] = 0x08; result[1] = 0xD6;
        result[2] = (byte)(length - 3); result[3] = (byte)((length - 3) >> 8);
        result[4] = (byte)count;
        for (int index = 0; index < edges.Length; index++) {
            result[5 + index * 2] = (byte)edges[index];
            result[6 + index * 2] = (byte)(edges[index] >> 8);
        }
        Buffer.BlockCopy(cells, 0, result, 5 + edges.Length * 2, cells.Length);
        return result;
    }
}
