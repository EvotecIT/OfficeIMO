using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsColorTests {
    private const string ProfileType = "application/vnd.ms-color.iccprofile";
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus");

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void ContextColorsRenderIndependentRgbAndCmykSwatches(XpsFormat format) {
        foreach (string row in File.ReadAllLines(Path.Combine(Corpus, "reference-srgb.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            var doc = XpsDocument.Create(format);
            string uri = doc.AddResource("Profiles/color.icc", File.ReadAllBytes(Path.Combine(Corpus, fields[0])), ProfileType);
            var page = doc.AddPage(20, 20);
            string channels = string.Join(",", fields[1].Split(':').Select(s => (int.Parse(s, CultureInfo.InvariantCulture) / 255D).ToString("R", CultureInfo.InvariantCulture)));
            page.AddPath("M0,0H20V20H0Z", "ContextColor " + uri + " 1," + channels);
            byte[] package = doc.Save();
            Assert.Contains(uri, Relationships(package));
            OfficeColor color = Raster(XpsDocument.Load(package).Pages[0]).GetPixel(10, 10);
            int[] expected = fields[2].Split(':').Select(s => int.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            Assert.InRange(Math.Abs(color.R - expected[0]), 0, 2);
            Assert.InRange(Math.Abs(color.G - expected[1]), 0, 2);
            Assert.InRange(Math.Abs(color.B - expected[2]), 0, 2);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExternalSolidAndGradientResourcesResolveTheirOwnProfileBase(bool gradient) {
        var doc = XpsDocument.Create(XpsFormat.Xps); var page = doc.AddPage(20, 20);
        string uri = doc.AddResource("Resources/Profiles/color.icc", File.ReadAllBytes(Path.Combine(Corpus, "littlecms-rgb-matrix.icc")), ProfileType);
        var xml = page.GetMarkup(); XNamespace ns = xml.Name.Namespace, key = "http://schemas.microsoft.com/winfx/2006/xaml";
        string color = "ContextColor ../Profiles/color.icc 2,2,-1,0";
        var brush = gradient ? new XElement(ns + "LinearGradientBrush", new XAttribute("StartPoint", "0,0"), new XAttribute("EndPoint", "20,0"),
            new XElement(ns + "LinearGradientBrush.GradientStops", new XElement(ns + "GradientStop", new XAttribute("Offset", "0"), new XAttribute("Color", color)),
                new XElement(ns + "GradientStop", new XAttribute("Offset", "1"), new XAttribute("Color", color))))
            : new XElement(ns + "SolidColorBrush", new XAttribute("Color", color));
        brush.SetAttributeValue(key + "Key", "paint");
        var dictionary = new XElement(ns + "ResourceDictionary", brush);
        string source = doc.AddResource("Resources/Dictionaries/paint.xaml", Encoding.UTF8.GetBytes(dictionary.ToString()), "application/vnd.ms-package.xps-resourcedictionary+xml");
        xml.Add(new XElement(ns + "FixedPage.Resources", new XElement(ns + "ResourceDictionary", new XAttribute("Source", source))),
            new XElement(ns + "Path", new XAttribute("Data", "M0,0H20V20H0Z"), new XAttribute("Fill", "{StaticResource paint}")));
        page.ReplaceMarkup(xml);
        byte[] package = doc.Save();
        Assert.Contains(uri, Relationships(package));
        var pixel = Raster(XpsDocument.Load(package).Pages[0]).GetPixel(10, 10);
        Assert.True(pixel.R >= 253 && pixel.G <= 2 && pixel.B <= 2 && pixel.A == 255);
        Assert.NotEmpty(doc.ToPdf());
    }

    [Fact]
    public void ProfileFailuresAndChannelMismatchCannotSilentlyRender() {
        var doc = XpsDocument.Create();
        string uri = doc.AddResource("Profiles/color.icc", new byte[128], ProfileType);
        var page = doc.AddPage(20, 20).AddPath("M0,0H20V20H0Z", "ContextColor " + uri + " 1,0,0,0");
        Assert.Throws<NotSupportedException>(() => page.ToSvg());
        Assert.Contains("Unsupported ICC profile", page.ToSvg(true).Diagnostics);
        doc.ReplaceResource(uri, File.ReadAllBytes(Path.Combine(Corpus, "littlecms-cmyk-lut.icc")));
        Assert.Throws<InvalidDataException>(() => page.ToSvg());
        doc.ReplaceResource(uri, new byte[4 * 1024 * 1024 + 1]);
        Assert.Throws<InvalidDataException>(() => page.ToSvg(true));
    }

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void TiffImagesPreservePixelsAndPhysicalViewbox(XpsFormat format) {
        var image = new OfficeRasterImage(8, 4, OfficeColor.Red);
        for (int y = 0; y < 4; y++) for (int x = 4; x < 8; x++) image.SetPixel(x, y, OfficeColor.Blue);
        byte[] tiff = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Tiff, new OfficeRasterEncodingOptions { DpiX = 192, DpiY = 192, WriteResolutionMetadata = true });
        var doc = XpsDocument.Create(format);
        string uri = doc.AddResource("Images/color.tif", tiff, "image/tiff");
        var page = doc.AddPage(80, 40).AddImage(uri, 0, 0, 80, 40);
        Assert.Equal("0,0,4,2", (string?)page.GetMarkup().Descendants().Single(e => e.Name.LocalName == "ImageBrush").Attribute("Viewbox"));
        page = XpsDocument.Load(doc.Save()).Pages[0];
        var raster = Raster(page);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(15, 20));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(65, 20));
        Assert.NotEmpty(doc.ToPdf());
    }

    [Fact]
    public void TiffWithEmbeddedProfileConvertsPixels() {
        byte[] tiff = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(2, 2, OfficeColor.Red), OfficeImageExportFormat.Tiff);
        Assert.Equal((byte)'I', tiff[0]);
        int oldIfd = BitConverter.ToInt32(tiff, 4);
        int count = BitConverter.ToUInt16(tiff, oldIfd);
        byte[] profile = File.ReadAllBytes(Path.Combine(Corpus, "icc-dci-p3-matrix.icc"));
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write(tiff); int newIfd = (int)stream.Position;
        writer.Write((ushort)(count + 1)); writer.Write(tiff, oldIfd + 2, count * 12);
        writer.Write((ushort)34675); writer.Write((ushort)7); writer.Write(profile.Length);
        writer.Write(newIfd + 2 + (count + 1) * 12 + 4); writer.Write(0); writer.Write(profile);
        stream.Position = 4; writer.Write(newIfd);
        var doc = XpsDocument.Create(); string uri = doc.AddResource("Images/profile.tif", stream.ToArray(), "image/tiff");
        var page = doc.AddPage(20, 20).AddImage(uri, 0, 0, 20, 20);
        Assert.Empty(page.ToSvg().Diagnostics);
        Assert.Equal(OfficeColor.Red, Raster(page).GetPixel(10, 10));
    }

    private static OfficeRasterImage Raster(XpsPage page) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var raster));
        return raster!;
    }
    private static string Relationships(byte[] package) {
        using var stream = new MemoryStream(package);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        var entry = zip.Entries.Single(e => e.FullName.EndsWith(".fpage.rels", StringComparison.Ordinal));
        using var reader = new StreamReader(entry.Open()); return reader.ReadToEnd();
    }
}
