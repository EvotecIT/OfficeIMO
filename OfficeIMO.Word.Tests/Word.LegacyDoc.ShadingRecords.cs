using OfficeIMO.Word.LegacyDoc.Model;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "E5EFF7")]
    [InlineData(true, "AUTO")]
    public void LegacyDoc_ModernParagraphShadingOverridesPaletteAndPreservesFollowingProperties(bool automatic, string fill) {
        byte[] record = {
            0x2D, 0x44, 0xC0, 0, // Older palette record: red.
            0x4D, 0xC6, 10, 0, 0, 0, 0xFF,
            automatic ? (byte)0 : (byte)0xE5, automatic ? (byte)0 : (byte)0xEF,
            automatic ? (byte)0 : (byte)0xF7, automatic ? (byte)0xFF : (byte)0, 0, 0,
            0x06, 0x24, 1 // Keep with next after the variable-length record.
        };
        LegacyDocParagraphFormat format = LegacyDocParagraphFormattingReader.ReadGrpprl(record, 0, record.Length);
        Assert.Equal(fill, format.ParagraphShading?.FillColorHex);
        Assert.True(format.KeepWithNext);
    }

    [Fact]
    public void LegacyDoc_TruncatedRgbShadingRetainsPriorFormattingWithoutReadingPastTheRecord() {
        byte[] record = { 0x2D, 0x44, 0xC0, 0, 0x4D, 0xC6, 10, 0, 0, 0, 0xFF, 0xE5, 0xEF, 0xF7, 0, 0 };
        LegacyDocParagraphFormat format = LegacyDocParagraphFormattingReader.ReadGrpprl(record, 0, record.Length);
        Assert.Equal("FF0000", format.ParagraphShading?.FillColorHex);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_AutomaticRunColorRecordsPreserveAutomaticAndFollowingProperties(bool modern) {
        byte[] record = modern
            ? new byte[] { 0x70, 0x68, 0x12, 0x34, 0x56, 0xFF, 0x35, 0x08, 1 }
            : new byte[] { 0x42, 0x2A, 0, 0x35, 0x08, 1 };
        LegacyDocCharacterFormat actual = LegacyDocCharacterFormattingReader.ReadGrpprl(record, 0, record.Length, Array.Empty<string>());
        Assert.Equal("auto", actual.ColorHex);
        Assert.True(actual.IsSpecified(LegacyDocCharacterFormatProperties.Color));
        Assert.True(actual.Bold);
    }
}
