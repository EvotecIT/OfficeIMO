using OfficeIMO.Word.Html;
using System;
using System.IO;
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
                var baseHref = new Uri(dir + Path.DirectorySeparatorChar).AbsoluteUri;
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
    }
}
