using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using OfficeIMO.Word.Html;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_InlineSvg_RoundTrip() {
            const string screenCss = ".shape{fill:blue}p{color:red}";
            const string printCss = ".shape{fill:green}";
            string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'>" +
                "<style media='screen'>" + screenCss + "</style><g><style media='print'>" + printCss +
                "</style><rect class='shape' width='10' height='10'/></g></svg>";
            string html = $"<style>@layer report{{p{{color:#123456}}}}</style><p>Outside SVG</p><p>{svg}</p>";
            using var doc = HtmlConversionDocument.Parse(html).WithTheme(OfficeTheme.Default).ToWord(new HtmlWordExportOptions());
            Assert.Single(doc.Images);
            using var bytes = new MemoryStream(doc.ToBytes(), false);
            using var package = WordprocessingDocument.Open(bytes, false);
            var main = package.MainDocumentPart!;
            using var image = Assert.Single(main.ImageParts, part => part.ContentType == "image/svg+xml").GetStream();
            XDocument savedSvg = XDocument.Load(image);
            XNamespace ns = "http://www.w3.org/2000/svg";
            XElement[] styles = savedSvg.Descendants(ns + "style").ToArray();
            Assert.Equal(new[] { screenCss, printCss }, styles.Select(style => style.Value));
            Assert.Equal(new[] { "screen", "print" }, styles.Select(style => (string?)style.Attribute("media")));
            Assert.Equal("svg", styles[0].Parent!.Name.LocalName);
            Assert.Equal(new[] { "style", "rect" }, styles[1].Parent!.Elements().Select(element => element.Name.LocalName));
            W.Run outside = Assert.Single(main.Document.Body!.Descendants<W.Run>(), run => run.InnerText == "Outside SVG");
            Assert.Equal("123456", outside.RunProperties!.Color!.Val!.Value);
            string back = doc.ToHtml();
            Assert.Contains("<svg", back, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_SvgImg_RoundTrip() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "Sample.svg");
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<p><img src=\"data:image/svg+xml;base64,{base64}\" width=\"10\" height=\"10\"></p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord(new HtmlWordExportOptions());
            Assert.Single(doc.Images);
            string back = doc.ToHtml();
            Assert.Contains("<svg", back, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_InlineSvgDimensionsUseInvariantHtmlNumbers() {
            CultureInfo previousCulture = CultureInfo.CurrentCulture;
            try {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
                const string html = "<svg xmlns='http://www.w3.org/2000/svg' width='10.5px' height='5.5px'><rect width='10' height='5'/></svg>";

                using var document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord();
                var image = Assert.Single(document.Images);

                Assert.InRange(image.Width!.Value.ToPixels(), 10.49D, 10.51D);
                Assert.InRange(image.Height!.Value.ToPixels(), 5.49D, 5.51D);
            } finally {
                CultureInfo.CurrentCulture = previousCulture;
            }
        }
    }
}
