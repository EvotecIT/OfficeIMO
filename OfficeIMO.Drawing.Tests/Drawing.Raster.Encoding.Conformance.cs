using System;
using System.IO;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class DrawingRasterEncodingTests {
    [Theory]
    [InlineData(1, "ABCDEF")]
    [InlineData(2, "CBAFED")]
    [InlineData(3, "FEDCBA")]
    [InlineData(4, "DEFABC")]
    [InlineData(5, "ADBECF")]
    [InlineData(6, "DAEBFC")]
    [InlineData(7, "FCEBDA")]
    [InlineData(8, "CFBEAD")]
    public void RasterLosslessExportsRetainAllEightOrientationsAndVisibleAlpha(int orientation, string order) {
        OfficeColor[] colors = {
            OfficeColor.FromRgba(240, 30, 40, 255), OfficeColor.FromRgba(20, 220, 50, 192),
            OfficeColor.FromRgba(30, 40, 210, 128), OfficeColor.FromRgba(10, 190, 220, 64),
            OfficeColor.FromRgba(210, 30, 180, 32), OfficeColor.FromRgba(230, 210, 20, 16)
        };
        var source = new OfficeRasterImage(3, 2);
        for (int index = 0; index < colors.Length; index++) source.SetPixel(index % 3, index / 3, colors[index]);
        byte[] tiff = OfficeRasterImageEncoder.Encode(source, OfficeImageExportFormat.Tiff);
        int entry = FindClassicTiffEntry(tiff, 274);
        WriteLittleEndian(tiff, entry + 8, orientation);
        Assert.True(OfficeRasterImageDecoder.TryDecode(tiff, out OfficeRasterImage? oriented));
        Assert.Equal(orientation >= 5 ? 2 : 3, oriented!.Width);
        Assert.Equal(orientation >= 5 ? 3 : 2, oriented.Height);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_RASTER_CONFORMANCE_ARTIFACTS");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output);
            File.WriteAllBytes(Path.Combine(output, "orientation-" + orientation + ".tiff"), tiff);
        }
        foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Png, OfficeImageExportFormat.Tiff, OfficeImageExportFormat.Webp }) {
            byte[] encoded = OfficeRasterImageEncoder.Encode(oriented, format);
            Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? decoded));
            Assert.Equal(oriented.Width, decoded!.Width);
            Assert.Equal(oriented.Height, decoded.Height);
            for (int index = 0; index < order.Length; index++) {
                Assert.Equal(colors[order[index] - 'A'], decoded.GetPixel(index % decoded.Width, index / decoded.Width));
            }
        }
    }
}
