using OfficeIMO.Word.Html;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_RelativeImage_UsesBaseUrl() {
            var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(dir);
            var source = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            var dest = Path.Combine(dir, "logo.png");
            File.Copy(source, dest);
            try {
                var baseHref = new Uri(new Uri(Path.Combine(dir, "dummy"), UriKind.Absolute), ".").AbsoluteUri;
                Assert.EndsWith("/", baseHref);
                string html = $"<base href=\"{baseHref}\"><img src=\"logo.png\" alt=\"Logo\" />";
                var doc = html.LoadFromHtml(new HtmlToWordOptions());
                Assert.Single(doc.Images);
                Assert.Equal("Logo", doc.Images[0].Description);
            } finally {
                File.Delete(dest);
                Directory.Delete(dir);
            }
        }

        [Fact]
        public void HtmlToWord_RelativeImage_UsesOptionsBasePath() {
            var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(dir);
            var source = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            var dest = Path.Combine(dir, "logo.png");
            File.Copy(source, dest);
            try {
                string html = "<img src=\"logo.png\" alt=\"Logo\" />";
                var options = new HtmlToWordOptions { BasePath = dir };
                var doc = html.LoadFromHtml(options);
                Assert.Single(doc.Images);
                Assert.Equal("Logo", doc.Images[0].Description);
            } finally {
                File.Delete(dest);
                Directory.Delete(dir);
            }
        }

        [Fact]
        public void HtmlToWord_UnreachableImage_InsertsPlaceholder() {
            string html = "<img src=\"http://localhost:1/missing.png\" alt=\"Missing\" />";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Missing", doc.Paragraphs[0].Text);
        }

        [Fact]
        public void HtmlToWord_UnreachableImage_NoAlt_SkipsImage() {
            string html = "<img src=\"http://localhost:1/missing.png\" />";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            Assert.Empty(doc.Images);
            Assert.Empty(doc.Paragraphs);
        }

        [Fact]
        public void InlineImagePreservesTextOrder() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<p>before<img src=\"data:image/png;base64,{base64}\">after</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            Assert.Equal(3, doc.Paragraphs.Count);
            Assert.Equal("before", doc.Paragraphs[0].Text);
            Assert.NotNull(doc.Paragraphs[1].Image);
            Assert.Equal("after", doc.Paragraphs[2].Text);
        }

        [Fact]
        public void DuplicateImageSrcSharesPart() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            var dataUri = $"data:image/png;base64,{base64}";
            string html = $"<p><img src=\"{dataUri}\"/><img src=\"{dataUri}\"/></p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            Assert.Collection(doc.Images, _ => { }, _ => { });
            Assert.Equal(doc.Images[0].RelationshipId, doc.Images[1].RelationshipId);
            Assert.Single(doc._wordprocessingDocument.MainDocumentPart.ImageParts);
        }
    }
}
