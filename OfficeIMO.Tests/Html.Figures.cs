using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void Html_FigureWithCaption_Converts() {
            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            byte[] imageBytes = File.ReadAllBytes(assetPath);
            string base64 = Convert.ToBase64String(imageBytes);
            string html = $"<figure><img src=\"data:image/png;base64,{base64}\" alt=\"Logo\"/><figcaption>Logo caption</figcaption></figure>";

            using var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Single(doc.Images);
            Assert.Equal("Logo caption", doc.Paragraphs[1].Text);
            Assert.Equal("Caption", doc.Paragraphs[1].StyleId);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<figure>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<figcaption>Logo caption</figcaption>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToHtml_FigureWithCaption_RendersFigure() {
            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            using var doc = WordDocument.Create();
            doc.AddParagraph().AddImage(assetPath);
            doc.AddParagraph("Logo caption").SetStyleId("Caption");

            string html = doc.ToHtml();
            Assert.Contains("<figure>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<figcaption>Logo caption</figcaption>", html, StringComparison.OrdinalIgnoreCase);
        }
    }
}
