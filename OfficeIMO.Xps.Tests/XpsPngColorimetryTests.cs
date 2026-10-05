using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsPngColorimetryTests {
    private static string Fixtures => Path.Combine(AppContext.BaseDirectory, "Fixtures", "PngColorimetry");

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void GammaAndChromaticityImagesUseNativeSrgbDefaultsAndRetainAlpha(XpsFormat format) {
        foreach (string row in File.ReadAllLines(Path.Combine(Fixtures, "expected.csv")).Skip(1)) {
            string[] values = row.Split(',');
            var document = Create(format, File.ReadAllBytes(Path.Combine(Fixtures, values[0])));
            var page = XpsDocument.Load(document.Save()).Pages[0];
            Assert.Empty(page.ToSvg().Diagnostics);
            OfficeColor pixel = Raster(page).GetPixel(8, 8);
            Assert.InRange(Math.Abs(pixel.R - int.Parse(values[3])), 0, 2);
            Assert.InRange(Math.Abs(pixel.G - int.Parse(values[4])), 0, 2);
            Assert.InRange(Math.Abs(pixel.B - int.Parse(values[5])), 0, 2);
            Assert.Equal(int.Parse(values[6]), (int)pixel.A);
            Assert.NotEmpty(document.ToPdf());
        }
    }

    [Fact]
    public void UsableAssociatedProfileOverridesNativePngDefaults() {
        var document = Create(XpsFormat.OpenXps, File.ReadAllBytes(Path.Combine(Fixtures, "gamma-linear.png")));
        string profile = document.AddResource("Profiles/source.icc", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus", "icc-dci-p3-matrix.icc")), "application/vnd.ms-color.iccprofile");
        var page = document.Pages[0]; var markup = page.GetMarkup();
        markup.Descendants().Single(e => e.Name.LocalName == "ImageBrush").SetAttributeValue("ImageSource", "{ColorConvertedBitmap /Images/source.png " + profile + "}");
        page.ReplaceMarkup(markup);
        var pixel = Raster(page).GetPixel(8, 8);
        Assert.InRange(pixel.R, 121, 125); Assert.InRange(pixel.G, 38, 42); Assert.InRange(pixel.B, 0, 2); Assert.Equal(128, pixel.A);
    }

    [Theory]
    [InlineData("missing-gamma")]
    [InlineData("singular-primaries")]
    public void NativeDefaultDoesNotDependOnPngTransferOrPrimaryCalibration(string kind) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Fixtures, "p3.png"));
        if (kind == "missing-gamma") bytes = RewriteChunk(bytes, "gAMA", null);
        else if (kind == "singular-primaries") {
            int[] coordinates = { 31270,32900,68000,32000,68000,32000,15000,6000 };
            bytes = RewriteChunk(bytes, "cHRM", coordinates.SelectMany(BigEndian).ToArray());
        }
        var page = Create(XpsFormat.Xps, bytes).Pages[0];
        var pixel = Raster(page).GetPixel(8, 8);
        Assert.Equal(128, pixel.R); Assert.Equal(64, pixel.G); Assert.Equal(32, pixel.B);
    }

    [Fact]
    public void UnsupportedCicpStillRequiresExplicitColorSupport() {
        byte[] bytes = RewriteChunk(File.ReadAllBytes(Path.Combine(Fixtures, "p3.png")), "cICP", new byte[] { 1, 16, 0, 1 });
        var page = Create(XpsFormat.Xps, bytes).Pages[0];
        Assert.Throws<NotSupportedException>(() => page.ToSvg());
        Assert.Contains("Image color metadata without a supported ICC profile", page.ToSvg(true).Diagnostics);
    }

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void PngResourcesUseStaticImageInsteadOfSeparateAnimationFrame(XpsFormat format) {
        foreach (string kind in new[] { "ordinary", "gamma", "associated", "embedded" }) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(Fixtures, kind == "embedded" ? "apng-static-embedded.png" : "apng-static.png"));
            if (kind == "gamma") bytes = RewriteChunk(bytes, "gAMA", BigEndian(100000));
            var document = Create(format, bytes);
            if (kind == "associated") {
                string profile = document.AddResource("Profiles/source.icc", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus", "littlecms-rgb-matrix.icc")), "application/vnd.ms-color.iccprofile");
                var markup = document.Pages[0].GetMarkup();
                markup.Descendants().Single(e => e.Name.LocalName == "ImageBrush").SetAttributeValue("ImageSource", "{ColorConvertedBitmap /Images/source.png " + profile + "}");
                document.Pages[0].ReplaceMarkup(markup);
            }
            var page = XpsDocument.Load(document.Save()).Pages[0];
            Assert.Empty(page.ToSvg().Diagnostics);
            OfficeColor pixel = Raster(page).GetPixel(8, 8);
            Assert.InRange(pixel.R, 250, 255); Assert.InRange(pixel.G, 0, 5); Assert.InRange(pixel.B, 0, 5); Assert.Equal(255, pixel.A);
            Assert.NotEmpty(document.ToPdf());
        }
    }

    private static XpsDocument Create(XpsFormat format, byte[] bytes) {
        var document = XpsDocument.Create(format);
        string source = document.AddResource("Images/source.png", bytes, "image/png");
        document.AddPage(16, 16).AddImage(source, 0, 0, 16, 16); return document;
    }
    private static OfficeRasterImage Raster(XpsPage page) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png, new OfficeImageExportOptions { BackgroundColor = OfficeColor.Transparent }).Bytes, out var raster));
        return raster!;
    }
    private static byte[] BigEndian(int value) => new[] { (byte)(value >> 24), (byte)(value >> 16), (byte)(value >> 8), (byte)value };
    private static byte[] RewriteChunk(byte[] source, string name, byte[]? payload) {
        using var output = new MemoryStream(); output.Write(source, 0, 8); bool inserted = false;
        for (int offset = 8; offset < source.Length;) {
            int length = (source[offset] << 24) | (source[offset + 1] << 16) | (source[offset + 2] << 8) | source[offset + 3];
            string type = System.Text.Encoding.ASCII.GetString(source, offset + 4, 4);
            if (!inserted && payload != null && (type == name || type == "IDAT")) {
                byte[] content = System.Text.Encoding.ASCII.GetBytes(name).Concat(payload).ToArray(); uint crc = uint.MaxValue;
                foreach (byte value in content) { crc ^= value; for (int bit = 0; bit < 8; bit++) crc = (crc >> 1) ^ ((crc & 1) != 0 ? 0xedb88320U : 0U); }
                byte[] chunk = BigEndian(payload.Length).Concat(content).Concat(BigEndian(unchecked((int)~crc))).ToArray(); output.Write(chunk, 0, chunk.Length); inserted = true;
            }
            if (type != name) output.Write(source, offset, length + 12);
            offset += length + 12;
        }
        return output.ToArray();
    }
}
