using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeArtPropertyTableReader_DecodesFixedAndComplexEntries() {
        byte[] payload = {
            0x81, 0x01, 0x33, 0x22, 0x11, 0x00,
            0x80, 0x83, 0x06, 0x00, 0x00, 0x00,
            0x4E, 0x00, 0x61, 0x00, 0x6D, 0x00
        };

        IReadOnlyList<OfficeArtProperty> properties = OfficeArtPropertyTableReader.Read(payload, 2);

        Assert.Equal(2, properties.Count);
        Assert.Equal("fillColor", properties[0].PropertyName);
        Assert.Equal("Fill", properties[0].PropertyGroupName);
        Assert.Equal(0x00112233U, properties[0].Value);
        Assert.True(properties[1].IsComplex);
        Assert.Equal("wzName", properties[1].PropertyName);
        Assert.Equal(6, properties[1].AvailableComplexDataLength);
        Assert.Equal("Nam", properties[1].ComplexText);
        Assert.Equal(new byte[] { 0x4E, 0x00, 0x61, 0x00, 0x6D, 0x00 },
            properties[1].CopyComplexData());
    }

    [Fact]
    public void OfficeArtShapeStyle_DecodesVisibilityColorsAndLineGeometry() {
        byte[] payload = {
            0x81, 0x01, 0x00, 0x00, 0x00, 0x08,
            0x82, 0x01, 0x00, 0x80, 0x00, 0x00,
            0xBF, 0x01, 0x10, 0x00, 0x10, 0x00,
            0xC0, 0x01, 0x33, 0x22, 0x11, 0x00,
            0xCB, 0x01, 0x00, 0x7F, 0x00, 0x00,
            0xCE, 0x01, 0x03, 0x00, 0x00, 0x00,
            0xFF, 0x01, 0x08, 0x00, 0x08, 0x00
        };

        OfficeArtShapeStyle style = OfficeArtShapeStyle.Decode(
            OfficeArtPropertyTableReader.Read(payload, 7));

        Assert.True(style.FillEnabled);
        Assert.Equal(0.5D, style.FillOpacity);
        Assert.True(style.LineEnabled);
        Assert.Equal(32512, style.LineWidthEmus);
        Assert.Equal(3U, style.LineDashing);
        Assert.True(style.FillColor!.Value.TryResolve(
            index => index == 0 ? OfficeColor.FromRgb(0xAA, 0xBB, 0xCC) : null,
            out OfficeColor fill));
        Assert.Equal(OfficeColor.FromRgb(0xAA, 0xBB, 0xCC), fill);
        Assert.True(style.LineColor!.Value.TryResolve(null, out OfficeColor line));
        Assert.Equal(OfficeColor.FromRgb(0x33, 0x22, 0x11), line);
    }

    [Fact]
    public void OfficeArtPropertyTableReader_RejectsTruncatedFixedTableWithoutOverread() {
        byte[] payload = { 0x81, 0x01, 0x33, 0x22, 0x11 };

        Assert.Empty(OfficeArtPropertyTableReader.Read(payload, 1));
    }
}
