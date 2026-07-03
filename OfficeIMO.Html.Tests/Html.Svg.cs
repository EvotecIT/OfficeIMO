using OfficeIMO.Word.Html;
using System;
using System.IO;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_InlineSvg_RoundTrip() {
            string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><rect width='10' height='10' fill='blue'/></svg>";
            string html = $"<p>{svg}</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            Assert.Single(doc.Images);
            string back = doc.ToHtml();
            Assert.Contains("<svg", back, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_SvgImg_RoundTrip() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "Sample.svg");
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<p><img src=\"data:image/svg+xml;base64,{base64}\" width=\"10\" height=\"10\"></p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            Assert.Single(doc.Images);
            string back = doc.ToHtml();
            Assert.Contains("<svg", back, StringComparison.OrdinalIgnoreCase);
        }
    }
}
