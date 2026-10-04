using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsImageColorTests {
    private static string Images => Path.Combine(AppContext.BaseDirectory, "Fixtures", "ColorImages");
    private static byte[] Profile(string name) => File.ReadAllBytes(name.StartsWith("gray-", StringComparison.Ordinal) ? Path.Combine(Images, name) : Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus", name));

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void AssociatedAndEmbeddedImagesMatchIndependentColorConversion(XpsFormat format) {
        foreach (string row in File.ReadAllLines(Path.Combine(Images, "expected.csv")).Skip(1)) {
            string[] values = row.Split(',');
            string file = values[0];
            var doc = XpsDocument.Create(format);
            string source = doc.AddResource("Images/source" + Path.GetExtension(file), File.ReadAllBytes(Path.Combine(Images, file)), Type(file));
            var page = doc.AddPage(16, 16).AddImage(source, 0, 0, 16, 16);
            if (file.Contains("associated")) {
                string profile = doc.AddResource("Profiles/source.icc", Profile(values[1]), "application/vnd.ms-color.iccprofile");
                Source(page, "{ColorConvertedBitmap " + source + " " + profile + "}");
                using var memory = new MemoryStream(doc.Save());
                using var zip = new ZipArchive(memory, ZipArchiveMode.Read);
                using var reader = new StreamReader(zip.Entries.Single(e => e.FullName.EndsWith(".fpage.rels", StringComparison.Ordinal)).Open());
                string relationships = reader.ReadToEnd();
                Assert.Contains(profile, relationships); Assert.Contains(source, relationships);
            }
            page = XpsDocument.Load(doc.Save()).Pages[0];
            var pixel = Raster(page).GetPixel(8, 8);
            Assert.InRange(Math.Abs(pixel.R - int.Parse(values[2])), 0, 2);
            Assert.InRange(Math.Abs(pixel.G - int.Parse(values[3])), 0, 2);
            Assert.InRange(Math.Abs(pixel.B - int.Parse(values[4])), 0, 2);
            Assert.Equal(int.Parse(values[5]), (int)pixel.A);
            Assert.NotEmpty(doc.ToPdf());
        }
    }

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void AssociatedProfileOverridesEmbeddedAndUsesDictionaryBase(XpsFormat format) {
        var doc = XpsDocument.Create(format);
        doc.AddResource("Resources/Images/source.png", File.ReadAllBytes(Path.Combine(Images, "RGB-png-embedded.png")), "image/png");
        doc.AddResource("Resources/Profiles/source.icc", Profile("littlecms-rgb-matrix.icc"), "application/vnd.ms-color.iccprofile");
        var page = doc.AddPage(16, 16); var markup = page.GetMarkup(); XNamespace ns = markup.Name.Namespace;
        XNamespace x = format == XpsFormat.Xps ? "http://schemas.microsoft.com/winfx/2006/xaml" : "http://schemas.openxps.org/oxps/v1.0/resourcedictionary-key";
        var dictionary = new XElement(ns + "ResourceDictionary", new XElement(ns + "ImageBrush", new XAttribute(x + "Key", "paint"),
            new XAttribute("ImageSource", "{ColorConvertedBitmap ../Images/source.png ../Profiles/source.icc}"),
            new XAttribute("Viewbox", "0,0,16,16"), new XAttribute("Viewport", "0,0,16,16")));
        string resource = doc.AddResource("Resources/Dictionaries/brush.xaml", System.Text.Encoding.UTF8.GetBytes(dictionary.ToString()), "application/vnd.ms-package.xps-resourcedictionary+xml");
        markup.Add(new XElement(ns + "FixedPage.Resources", new XElement(ns + "ResourceDictionary", new XAttribute("Source", resource))),
            new XElement(ns + "Path", new XAttribute("Data", "M0,0H16V16H0Z"), new XAttribute("Fill", "{StaticResource paint}")));
        page.ReplaceMarkup(markup);
        var pixel = Raster(XpsDocument.Load(doc.Save()).Pages[0]).GetPixel(8, 8);
        Assert.InRange(pixel.R, 127, 129); Assert.InRange(pixel.G, 63, 65); Assert.InRange(pixel.B, 31, 33);
    }

    [Fact]
    public void ImageProfileFailuresAreReportedAndCancellationIsObserved() {
        var doc = XpsDocument.Create();
        string source = doc.AddResource("Images/source.png", File.ReadAllBytes(Path.Combine(Images, "RGB-png-associated.png")), "image/png");
        string profile = doc.AddResource("Profiles/source.icc", new byte[128], "application/vnd.ms-color.iccprofile");
        var page = doc.AddPage(16, 16).AddImage(source, 0, 0, 16, 16);
        Source(page, "{ColorConvertedBitmap " + source + " " + profile + "}");
        Assert.Throws<NotSupportedException>(() => page.ToSvg());
        Assert.Contains("Unsupported ICC profile", page.ToSvg(true).Diagnostics);
        doc.ReplaceResource(profile, Profile("littlecms-cmyk-lut.icc"));
        Assert.Throws<NotSupportedException>(() => page.ToSvg());
        doc.ReplaceResource(profile, Profile("icc-dci-p3-matrix.icc"));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => page.ToSvg(cancellationToken: cancellation.Token));
    }

    [Theory]
    [InlineData("oversized-profile.png")]
    [InlineData("RGB-jpg-embedded.jpg")]
    public void OversizedAndIncompleteEmbeddedProfilesDoNotSilentlyLoseColor(string file) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Images, file));
        if (file.EndsWith(".jpg", StringComparison.Ordinal)) {
            byte[] marker = System.Text.Encoding.ASCII.GetBytes("ICC_PROFILE\0");
            int start = Enumerable.Range(0, bytes.Length - marker.Length).First(i => bytes.Skip(i).Take(marker.Length).SequenceEqual(marker));
            bytes[start + marker.Length + 1] = 2; // Declare an absent second APP2 segment.
        }
        var doc = XpsDocument.Create(); string uri = doc.AddResource("Images/source" + Path.GetExtension(file), bytes, Type(file));
        var page = doc.AddPage(16, 16).AddImage(uri, 0, 0, 16, 16);
        Assert.Throws<NotSupportedException>(() => page.ToSvg());
        Assert.Contains("Unusable embedded ICC profile", page.ToSvg(true).Diagnostics);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ColorManagementPreservesJpegOrientation(bool progressive, bool embedded) {
        var doc = XpsDocument.Create();
        string source = doc.AddResource("Images/source.jpg", File.ReadAllBytes(Path.Combine(Images, $"oriented-{progressive}-{embedded}.jpg")), "image/jpeg");
        var page = doc.AddPage(32, 16).AddImage(source, 0, 0, 32, 16);
        if (!embedded) {
            // Ordinary JPEG conversion already honors EXIF; activating a profile must retain it.
            Assert.True(Raster(page).GetPixel(8, 8).B > 240);
            string profile = doc.AddResource("Profiles/source.icc", Profile("littlecms-rgb-matrix.icc"), "application/vnd.ms-color.iccprofile");
            Source(page, "{ColorConvertedBitmap " + source + " " + profile + "}");
        }
        var raster = Raster(page);
        Assert.True(raster.GetPixel(8, 8).B > 240);
        Assert.True(raster.GetPixel(24, 8).R > 240);
    }

    private static string Type(string name) => name.EndsWith(".png", StringComparison.Ordinal) ? "image/png" : name.EndsWith(".jpg", StringComparison.Ordinal) ? "image/jpeg" : "image/tiff";
    private static void Source(XpsPage page, string value) {
        var markup = page.GetMarkup(); markup.Descendants().Single(e => e.Name.LocalName == "ImageBrush").SetAttributeValue("ImageSource", value); page.ReplaceMarkup(markup);
    }
    private static OfficeRasterImage Raster(XpsPage page) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png, new OfficeImageExportOptions { BackgroundColor = OfficeColor.Transparent }).Bytes, out var image)); return image!;
    }
}
