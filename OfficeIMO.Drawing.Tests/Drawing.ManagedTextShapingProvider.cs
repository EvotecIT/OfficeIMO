using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingManagedTextShapingProviderTests {
    [Fact]
    public void ManagedProvider_ShapesSupportedArabicAndPreservesLogicalMappings() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(
            0x0627,
            0x0628,
            0xFE8D,
            0xFE8F);
        var request = new OfficeTextShapingRequest(
            "اب",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.RightToLeft,
            language: "ar");

        OfficeTextShapingResult? result = OfficeManagedTextShapingProvider.Instance.ShapeText(request);

        Assert.NotNull(result);
        Assert.Equal(2, result!.Glyphs.Count);
        Assert.Equal("ب", result.Glyphs[0].UnicodeText);
        Assert.Equal(1, result.Glyphs[0].TextIndex);
        Assert.Equal("ا", result.Glyphs[1].UnicodeText);
        Assert.Equal(0, result.Glyphs[1].TextIndex);
    }

    [Fact]
    public void ManagedProvider_DeclinesScriptsOutsideItsBoundedSubset() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(0x0915, 0x093F);
        var request = new OfficeTextShapingRequest(
            "कि",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000);

        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(request));
    }

    [Fact]
    public void ManagedProvider_PreservesMappingsWhenVisualGlyphsRepeat() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(0x0627, 0xFE8D);
        var request = new OfficeTextShapingRequest(
            "اا",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.RightToLeft,
            language: "ar");

        OfficeTextShapingResult? result = OfficeManagedTextShapingProvider.Instance.ShapeText(request);

        Assert.NotNull(result);
        Assert.Equal(new[] { 1, 0 }, result!.Glyphs.Select(static glyph => glyph.TextIndex));
    }

    [Fact]
    public void ManagedProvider_HonorsExplicitBaseDirectionForMixedText() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(
            ' ',
            'a',
            'b',
            'c',
            0x0627,
            0x0628,
            0xFE8D,
            0xFE8F);
        var request = new OfficeTextShapingRequest(
            "abc اب",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.RightToLeft,
            language: "ar");

        OfficeTextShapingResult? result = OfficeManagedTextShapingProvider.Instance.ShapeText(request);

        Assert.NotNull(result);
        Assert.Equal(new[] { 5, 4, 0, 1, 2, 3 }, result!.Glyphs.Select(static glyph => glyph.TextIndex));
    }

    [Fact]
    public void ManagedProvider_DeclinesExplicitBidiControlsItDoesNotImplement() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(0x61, 0x62, 0x63);
        var request = new OfficeTextShapingRequest(
            "\u202Eabc\u202C",
            ManagedTextShapingTestAssets.FamilyName,
            font,
            isOpenTypeCff: false,
            unitsPerEm: 1000,
            direction: OfficeTextDirection.RightToLeft);

        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(request));
    }
}
