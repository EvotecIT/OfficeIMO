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

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Single(doc.Images);
            Assert.Equal("Logo caption", doc.Paragraphs[1].Text);
            Assert.Equal("Caption", doc.Paragraphs[1].StyleId);
        }
    }
}
