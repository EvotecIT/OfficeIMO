using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsJpegXrExtendedColorTests {
    [Theory]
    [InlineData("s16-3c-frequency-q32-alpha0", 3, false)]
    [InlineData("s32-3c-frequency-q32-alpha0", 3, false)]
    [InlineData("f16-3c-frequency-q32-alpha0", 3, false)]
    [InlineData("f32-3c-frequency-q32-alpha0", 3, false)]
    [InlineData("f16-1c-frequency-q32-alpha0", 1, false)]
    [InlineData("f32-premultiplied-4c-frequency-q32-alpha2", 4, true)]
    public void ProfilesReceiveSourcePrecisionBeforeDefaultScRgbConversion(string id, int components, bool premultiplied) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Fixtures", "JpegXr");
        byte[] encoded = File.ReadAllBytes(Path.Combine(corpus, "extended-" + id + ".jxr"));
        byte[] profileBytes = File.ReadAllBytes(components == 1
            ? Path.Combine(AppContext.BaseDirectory, "Fixtures", "ColorImages", "gray-gamma18.icc")
            : Path.Combine(AppContext.BaseDirectory, "Fixtures", "IccColorCorpus", "littlecms-rgb-matrix.icc"));
        Assert.True(OfficeIccColorProfile.TryCreate(profileBytes, out var profile));
        var expected = new byte[19 * 13 * 4];
        using (var raw = new BinaryReader(File.OpenRead(Path.Combine(corpus, "extended-" + id + ".linear")))) {
            var values = new double[components];
            var colorChannels = new double[components == 1 ? 1 : 3];
            for (int pixel = 0; pixel < 19 * 13; pixel++) {
                for (int c = 0; c < components; c++) values[c] = raw.ReadDouble();
                double alpha = components == 4 ? values[3] : 1D;
                for (int c = 0; c < colorChannels.Length; c++)
                    colorChannels[c] = premultiplied ? alpha <= 0D ? 0D : values[c] / alpha : values[c];
                Assert.True(profile!.TryConvert(colorChannels, OfficeIccRenderingIntent.RelativeColorimetric, out var color));
                expected[pixel * 4] = color.R; expected[pixel * 4 + 1] = color.G; expected[pixel * 4 + 2] = color.B;
                expected[pixel * 4 + 3] = (byte)Math.Round(Math.Max(0D, Math.Min(1D, alpha)) * 255D);
            }
            Assert.Equal(raw.BaseStream.Length, raw.BaseStream.Position);
        }
        foreach (bool embedded in new[] { false, true }) {
            byte[] source = embedded ? OfficeIMO.TestAssets.JpegXrTestFixture.WithField(encoded, 0x8773, 7, profileBytes) : encoded;
            var document = XpsDocument.Create(XpsFormat.OpenXps);
            string imageUri = document.AddResource("Images/source.jxr", source, "image/jxr");
            var page = document.AddPage(19, 13).AddImage(imageUri, 0, 0, 19, 13);
            if (!embedded) {
                string profileUri = document.AddResource("Profiles/source.icc", profileBytes, "application/vnd.ms-color.iccprofile");
                var markup = page.GetMarkup();
                markup.Descendants().Single(e => e.Name.LocalName == "ImageBrush").SetAttributeValue("ImageSource",
                    "{ColorConvertedBitmap " + imageUri + " " + profileUri + "}");
                page.ReplaceMarkup(markup);
            }
            var svg = XpsDocument.Load(document.Save()).Pages[0].ToSvg();
            Assert.Empty(svg.Diagnostics);
            var markupSvg = XDocument.Parse(svg.Svg);
            string data = (string)markupSvg.Descendants().Single(e => e.Name.LocalName == "image").Attribute("href")!;
            Assert.True(OfficeRasterImageDecoder.TryDecode(Convert.FromBase64String(data.Substring(data.IndexOf(',') + 1)), out var pixels));
            Assert.Equal(expected, pixels!.GetPixels());
        }
    }
}
