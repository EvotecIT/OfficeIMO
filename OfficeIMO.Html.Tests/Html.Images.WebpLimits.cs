using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Word.Html;
using Xunit;
namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_WebpExifOrientationIsAppliedToEmbeddedPng() {
            const string source = "UklGRmQAAABXRUJQVlA4WAoAAAAIAAAAAgAAAQAAVlA4TCQAAAAvAkAAAC8gEEjaH3qN+RcQFPk/2vwHH0QCg0CADPHiSET/IxZFWElGGgAAAE1NACoAAAAIAAEBEgADAAAAAQAGAAAAAAAA";
            var result = HtmlConversionDocument.Parse("<img src='data:image/webp;base64," + source + "'>").ToWordDocumentResult();
            using var doc = result.Value;
            using var packageBytes = new MemoryStream(doc.ToBytes());
            using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(packageBytes, false);
            using var png = Assert.Single(package.MainDocumentPart!.ImageParts).GetStream();
            using var bytes = new MemoryStream();
            png.CopyTo(bytes);
            Assert.True(OfficeIMO.Drawing.OfficeRasterImageDecoder.TryDecode(bytes.ToArray(), out var image));
            Assert.Equal(2, image!.Width);
            Assert.Equal(3, image.Height);
            Assert.Equal(new byte[] { 255, 255, 0, 255 }, image.GetPixels().Take(4).ToArray());
        }

        [Theory]
        [InlineData("data", false)]
        [InlineData("file", false)]
        [InlineData("remote", false)]
        [InlineData("data", true)]
        [InlineData("file", true)]
        [InlineData("remote", true)]
        public async Task HtmlToWord_NormalizedRasterPixelLimitAppliesBeforeNormalization(string route, bool avif) {
            byte[] bytes = Convert.FromBase64String("UklGRjwAAABXRUJQVlA4IDAAAADQAQCdASoQABAAAUAmJaACdLoB+AADsAD+8ut//NgVzXPv9//S4P0uD9Lg/9KQAAA=");
            // Frozen 49x33 AVIF from the independent static-image corpus.
            if (avif) bytes = Convert.FromBase64String("AAAAIGZ0eXBhdmlmAAAAAGF2aWZtaWYxbWlhZk1BMUIAAADrbWV0YQAAAAAAAAAhaGRscgAAAAAAAAAAcGljdAAAAAAAAAAAAAAAAAAAAAAOcGl0bQAAAAAAAQAAAB5pbG9jAAAAAEQAAAEAAQAAAAEAAAETAAAAOgAAAChpaW5mAAAAAAABAAAAGmluZmUCAAAAAAEAAGF2MDFDb2xvcgAAAABqaXBycAAAAEtpcGNvAAAAFGlzcGUAAAAAAAAAMQAAACEAAAAQcGl4aQAAAAADCAgIAAAADGF2MUOBAAwAAAAAE2NvbHJuY2x4AAEADQAGgAAAABdpcG1hAAAAAAAAAAEAAQQBAoMEAAAAQm1kYXQSAAoJGBVwgaICGg0IMitEgf39qkAggg+QAADPzBAOzmhbr/5Sjc9uvYEuL65Q9kjfQOnL0iQ1J5+x");
            int pixels = avif ? 49 * 33 : 256;
            string extension = avif ? "avif" : "webp";
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "." + extension);
            using var client = new HttpClient(new FakeHtmlHttpMessageHandler(_ => Task.FromResult(
                new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(bytes) })));
            var options = HtmlToWordOptions.CreateTrustedDocumentProfile();
            options.MaxDecodedImagePixels = pixels - 1;
            options.HttpClient = client;
            File.WriteAllBytes(path, bytes);
            try {
                string source = route == "data" ? "data:image/" + extension + ";base64," + Convert.ToBase64String(bytes)
                    : route == "file" ? new Uri(path).AbsoluteUri : "https://example.test/image." + extension;
                var conversion = await HtmlConversionDocument.Parse("<img src='" + source + "' alt='Retained description'>",
                    HtmlConversionDocumentOptions.CreateTrustedProfile()).ToWordDocumentResultAsync(options);
                using var doc = conversion.Value;
                Assert.Empty(doc.Images);
                Assert.Contains(doc.Paragraphs, p => p.Text.Contains("Retained description"));
                Assert.Contains(conversion.Report.Diagnostics, d => (d.Detail ?? "").Contains("decoded-pixel"));
                options.MaxDecodedImagePixels = pixels;
                using var accepted = await HtmlConversionDocument.Parse("<img src='" + source + "'>",
                    HtmlConversionDocumentOptions.CreateTrustedProfile()).ToWordDocumentAsync(options);
                Assert.Single(accepted.Images);
            } finally { File.Delete(path); }
        }

        [Fact]
        public void HtmlToWord_UntrustedProfileBoundsHighlyCompressedWebp() {
            // libwebp lossless 2001 x 2000 solid red: small encoded input, more than four million decoded pixels.
            const string source = "UklGRsoAAABXRUJQVlA4TL4AAAAv0MfzAQcQ/Y/+BwQkSf//kxH9z/jPf/7zn//85z//+c9//vOf//znP//5z3/+85///Oc///nPf/7zn//85z//+c9//vOf//znP//5z3/+85///Oc///nPf/7zn//85z//+c9//vOf//znP//5z3/+85///Oc///nPf/7zn//85z//+c9//vOf//znP//5z3/+85///Oc///nPf/7zn//85z//+c9//vOf//znP//5z3/+85///Oc///nP/9EC";
            var result = HtmlConversionDocument.Parse("<img src='data:image/webp;base64," + source + "' alt='Large photo'>")
                .ToWordDocumentResult(HtmlToWordOptions.CreateUntrustedHtmlProfile());
            using var doc = result.Value;
            Assert.Empty(doc.Images);
            Assert.Contains(result.Report.Diagnostics, d => (d.Detail ?? "").Contains("decoded-pixel"));
        }
    }
}
