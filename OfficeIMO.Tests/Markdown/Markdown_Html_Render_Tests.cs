using System;
using System.IO;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Html_Render_Tests {
        [Fact]
        public void Fragment_Has_No_Html_Tag_And_Wrapper() {
            var doc = MarkdownDoc.Create().H1("Title").P("Hello");
            var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Clean });
            Assert.DoesNotContain("<html", html);
            Assert.Contains("<article", html);
            Assert.Contains("markdown-body", html);
        }

        [Fact]
        public void Document_Has_Full_Structure() {
            var doc = MarkdownDoc.Create().H1("Title").P("Hello");
            var html = doc.ToHtmlDocument(new HtmlOptions { Title = "Doc", Style = HtmlStyle.Clean });
            Assert.Contains("<html", html);
            Assert.Contains("<head>", html);
            Assert.Contains("</body>", html);
            Assert.Contains("<article", html);
        }

        [Fact]
        public void External_Css_Writes_Sidecar() {
            var temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            var path = Path.Combine(temp, "sample.html");
            var doc = MarkdownDoc.Create().H1("Title");
            doc.SaveHtml(path, new HtmlOptions { Title = "Doc", Style = HtmlStyle.Clean, CssDelivery = CssDelivery.ExternalFile });
            Assert.True(File.Exists(path));
            var cssPath = Path.ChangeExtension(path, ".css");
            Assert.True(File.Exists(cssPath));
        }

        [Fact]
        public void Cdn_Online_Uses_Link_Tag() {
            var doc = MarkdownDoc.Create().H1("Title");
            var html = doc.ToHtmlDocument(new HtmlOptions {
                CssDelivery = CssDelivery.LinkHref,
                CssHref = "https://cdn.example.com/style.css",
                AssetMode = AssetMode.Online,
                BodyClass = "markdown-body"
            });
            Assert.Contains("<link", html);
            Assert.Contains("cdn.example.com", html);
        }

        [Fact]
        public void Cdn_Offline_Inlines_Style() {
            var doc = MarkdownDoc.Create().H1("Title");
            var html = doc.ToHtmlDocument(new HtmlOptions {
                CssDelivery = CssDelivery.LinkHref,
                CssHref = "https://cdn.example.com/style.css",
                AssetMode = AssetMode.Offline,
                BodyClass = "markdown-body"
            });
            Assert.DoesNotContain("<link", html);
            Assert.Contains("<style", html);
        }

        [Fact]
        public void Null_CssScopeSelector_DisablesArticleScope() {
            var doc = MarkdownDoc.Create().H1("Title").P("Hello world");
            var html = doc.ToHtmlDocument(new HtmlOptions {
                Style = HtmlStyle.Clean,
                BodyClass = null,
                CssScopeSelector = null
            });
            Assert.DoesNotContain("article.markdown-body", html, StringComparison.Ordinal);
            Assert.Contains("body h1", html, StringComparison.Ordinal);
            Assert.Contains("<body>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Prism_ManifestOnly_Does_Not_Emit_Tags() {
            var doc = MarkdownDoc.Create().H1("Code").Code("csharp", "Console.WriteLine(\"x\");");
            var parts = doc.ToHtmlParts(new HtmlOptions {
                EmitMode = AssetEmitMode.ManifestOnly,
                Prism = new PrismOptions { Enabled = true, Languages = { "csharp" }, Theme = PrismTheme.GithubAuto }
            });
            Assert.NotEmpty(parts.Assets);
            Assert.DoesNotContain("<script", parts.Head);
            Assert.DoesNotContain("<link", parts.Head);
        }

        [Fact]
        public void AssetMerger_Dedupes_And_Produces_Media() {
            var opts = new HtmlOptions { EmitMode = AssetEmitMode.ManifestOnly, Prism = new PrismOptions { Enabled = true, Theme = PrismTheme.GithubAuto, Languages = { "csharp" } } };
            var p1 = MarkdownDoc.Create().Code("csharp", "Console.WriteLine(1);").ToHtmlParts(opts);
            var p2 = MarkdownDoc.Create().Code("csharp", "Console.WriteLine(2);").ToHtmlParts(opts);
            var merged = OfficeIMO.Markdown.HtmlAssetMerger.Build(new[] { p1.Assets, p2.Assets });
            // Only one core and one light/dark theme pair expected
            Assert.Contains("data-asset-id=\"prism-core\"", merged.headLinks);
            Assert.Contains("prefers-color-scheme: dark", merged.headLinks);
        }

        [Fact]
        public void HorizontalRule_Renders_In_Markdown_And_Html() {
            var doc = MarkdownDoc.Create()
                .P("Before the break.")
                .Hr()
                .P("After the break.");

            string markdown = doc.ToMarkdown();
            Assert.Contains("\n---\n", markdown, StringComparison.Ordinal);

            string html = doc.ToHtmlFragment();
            Assert.Contains("<hr />", html, StringComparison.Ordinal);
        }
    }
}
