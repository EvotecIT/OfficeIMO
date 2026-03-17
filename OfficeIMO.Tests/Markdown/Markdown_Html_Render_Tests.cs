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
            string normalizedMarkdown = markdown.Replace("\r\n", "\n");
            Assert.Contains("\n---\n", normalizedMarkdown, StringComparison.Ordinal);

            string html = doc.ToHtmlFragment();
            Assert.Contains("<hr />", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Image_With_Dimensions_Renders_Size_And_Caption() {
            var doc = MarkdownDoc.Create()
                .Image("images/photo.png", alt: "Alt text", title: "Title text", width: 640, height: 480)
                .Caption("Photo caption");

            var block = Assert.IsType<ImageBlock>(Assert.Single(doc.Blocks));
            Assert.Equal(640d, block.Width.GetValueOrDefault());
            Assert.Equal(480d, block.Height.GetValueOrDefault());
            Assert.Equal("Photo caption", block.Caption);

            string markdown = doc.ToMarkdown().Replace("\r\n", "\n");
            Assert.Contains("![Alt text](images/photo.png \"Title text\"){width=640 height=480}", markdown, StringComparison.Ordinal);
            Assert.Contains("_Photo caption_", markdown, StringComparison.Ordinal);

            string html = doc.ToHtmlFragment();
            Assert.Contains("<img src=\"images/photo.png\" alt=\"Alt text\" title=\"Title text\" width=\"640\" height=\"480\" />", html, StringComparison.Ordinal);
            Assert.Contains("<div class=\"caption\">Photo caption</div>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Linked_Image_Block_Renders_Linked_Markdown_And_Html() {
            var doc = MarkdownDoc.Create()
                .Add(new ImageBlock("images/photo.png", "Alt text", "Title text", linkUrl: "https://example.com/docs", linkTitle: "Read more", linkTarget: "_self", linkRel: "nofollow"))
                .Caption("Photo caption");

            var block = Assert.IsType<ImageBlock>(Assert.Single(doc.Blocks));
            Assert.Equal("https://example.com/docs", block.LinkUrl);
            Assert.Equal("Read more", block.LinkTitle);
            Assert.Equal("_self", block.LinkTarget);
            Assert.Equal("nofollow", block.LinkRel);

            string markdown = doc.ToMarkdown().Replace("\r\n", "\n");
            Assert.Contains("[![Alt text](images/photo.png \"Title text\")](https://example.com/docs \"Read more\")", markdown, StringComparison.Ordinal);
            Assert.Contains("_Photo caption_", markdown, StringComparison.Ordinal);

            string html = doc.ToHtmlFragment();
            Assert.Contains("<a href=\"https://example.com/docs\" title=\"Read more\" target=\"_self\" rel=\"nofollow\"", html, StringComparison.Ordinal);
            Assert.Contains("<img src=\"images/photo.png\" alt=\"Alt text\" title=\"Title text\"", html, StringComparison.Ordinal);
            Assert.Contains("<div class=\"caption\">Photo caption</div>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Linked_Image_Block_Hardens_TargetBlank_Rel_In_Html() {
            var doc = MarkdownDoc.Create()
                .Add(new ImageBlock("images/photo.png", "Alt text", "Title text", linkUrl: "https://example.com/docs", linkTitle: "Read more", linkTarget: "_blank"))
                .Caption("Photo caption");

            var block = Assert.IsType<ImageBlock>(Assert.Single(doc.Blocks));
            Assert.Equal("_blank", block.LinkTarget);
            Assert.Null(block.LinkRel);

            string markdown = doc.ToMarkdown().Replace("\r\n", "\n");
            Assert.Contains("[![Alt text](images/photo.png \"Title text\")](https://example.com/docs \"Read more\")", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("noopener", markdown, StringComparison.Ordinal);

            string html = doc.ToHtmlFragment();
            Assert.Contains("<a href=\"https://example.com/docs\" title=\"Read more\" target=\"_blank\" rel=\"", html, StringComparison.Ordinal);
            Assert.Contains("noopener", html, StringComparison.Ordinal);
            Assert.Contains("noreferrer", html, StringComparison.Ordinal);
            Assert.Contains("<img src=\"images/photo.png\" alt=\"Alt text\" title=\"Title text\"", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Inline_Link_Preserves_Html_Metadata_And_Hardens_TargetBlank() {
            var doc = MarkdownDoc.Create()
                .Add(new ParagraphBlock(new InlineSequence().Link("Read more", "https://example.com/docs", "Hero docs", "_blank", "nofollow sponsored")));

            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(doc.Blocks));
            var link = Assert.IsType<LinkInline>(Assert.Single(paragraph.Inlines.Nodes));
            Assert.Equal("_blank", link.LinkTarget);
            Assert.Equal("nofollow sponsored", link.LinkRel);

            string markdown = doc.ToMarkdown().Replace("\r\n", "\n");
            Assert.Contains("[Read more](https://example.com/docs \"Hero docs\")", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("noopener", markdown, StringComparison.Ordinal);

            string html = doc.ToHtmlFragment();
            Assert.Contains("<a href=\"https://example.com/docs\" title=\"Hero docs\" target=\"_blank\" rel=\"", html, StringComparison.Ordinal);
            Assert.Contains("nofollow", html, StringComparison.Ordinal);
            Assert.Contains("sponsored", html, StringComparison.Ordinal);
            Assert.Contains("noopener", html, StringComparison.Ordinal);
            Assert.Contains("noreferrer", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Image_Block_Renders_Picture_Sources_Back_To_Html() {
            var image = new ImageBlock("https://example.com/media/storm-wide.webp", "Flooded street at dawn", "Open full photo", 1280, 720) {
                PictureFallbackPath = "https://example.com/media/storm-fallback.jpg"
            };
            image.PictureSources.Add(new ImagePictureSource(
                "https://example.com/media/storm-wide.webp",
                "(min-width: 960px)",
                "image/webp",
                "100vw",
                "https://example.com/media/storm-wide.webp 1x, https://example.com/media/storm-wide@2x.webp 2x"));
            image.PictureSources.Add(new ImagePictureSource(
                "https://example.com/media/storm-mobile.webp",
                "(max-width: 959px)",
                "image/webp",
                "100vw",
                "https://example.com/media/storm-mobile.webp 1x, https://example.com/media/storm-mobile@2x.webp 2x"));

            var doc = MarkdownDoc.Create().Add(image).Caption("Residents navigate floodwater after the overnight storm.");

            string markdown = doc.ToMarkdown().Replace("\r\n", "\n");
            Assert.Contains("![Flooded street at dawn](https://example.com/media/storm-wide.webp \"Open full photo\")", markdown, StringComparison.Ordinal);

            string html = doc.ToHtmlFragment();
            Assert.Contains("<picture>", html, StringComparison.Ordinal);
            Assert.Contains("<source srcset=\"https://example.com/media/storm-wide.webp 1x, https://example.com/media/storm-wide@2x.webp 2x\" media=\"(min-width: 960px)\" type=\"image/webp\" sizes=\"100vw\" />", html, StringComparison.Ordinal);
            Assert.Contains("<source srcset=\"https://example.com/media/storm-mobile.webp 1x, https://example.com/media/storm-mobile@2x.webp 2x\" media=\"(max-width: 959px)\" type=\"image/webp\" sizes=\"100vw\" />", html, StringComparison.Ordinal);
            Assert.Contains("<img src=\"https://example.com/media/storm-fallback.jpg\" alt=\"Flooded street at dawn\" title=\"Open full photo\" width=\"1280\" height=\"720\"", html, StringComparison.Ordinal);
        }

        [Fact]
        public void Image_Block_Renders_Picture_Sources_With_Width_Descriptors_And_Query_Strings() {
            var image = new ImageBlock("https://example.com/media/storm-wide.webp?fit=cover&crop=10,20,300,400", "Flooded street at dawn", "Open full photo", 1280, 720) {
                PictureFallbackPath = "https://example.com/media/storm-fallback.jpg?download=1"
            };
            image.PictureSources.Add(new ImagePictureSource(
                "https://example.com/media/storm-wide.webp?fit=cover&crop=10,20,300,400",
                "(min-width: 960px)",
                "image/webp",
                "(min-width: 960px) 90vw, 100vw",
                "https://example.com/media/storm-wide.webp?fit=cover&crop=10,20,300,400 640w, https://example.com/media/storm-wide.webp?fit=cover&crop=20,40,600,800 1280w"));
            image.PictureSources.Add(new ImagePictureSource(
                "https://example.com/media/storm-mobile.webp?fit=cover&crop=5,10,200,250",
                "(max-width: 959px)",
                "image/webp",
                "100vw",
                "https://example.com/media/storm-mobile.webp?fit=cover&crop=5,10,200,250 320w, https://example.com/media/storm-mobile.webp?fit=cover&crop=10,20,400,500 640w"));

            string html = MarkdownDoc.Create().Add(image).ToHtmlFragment();
            Assert.Contains("<source srcset=\"https://example.com/media/storm-wide.webp?fit=cover&amp;crop=10,20,300,400 640w, https://example.com/media/storm-wide.webp?fit=cover&amp;crop=20,40,600,800 1280w\" media=\"(min-width: 960px)\" type=\"image/webp\" sizes=\"(min-width: 960px) 90vw, 100vw\" />", html, StringComparison.Ordinal);
            Assert.Contains("<source srcset=\"https://example.com/media/storm-mobile.webp?fit=cover&amp;crop=5,10,200,250 320w, https://example.com/media/storm-mobile.webp?fit=cover&amp;crop=10,20,400,500 640w\" media=\"(max-width: 959px)\" type=\"image/webp\" sizes=\"100vw\" />", html, StringComparison.Ordinal);
            Assert.DoesNotContain("%20", html, StringComparison.Ordinal);
        }

        [Fact]
        public void HtmlOptions_CodeBlockHtmlRenderer_Applies_Inside_Nested_List_Items() {
            var markdown = """
- item

  ```text
  hello from list
  ```
""";
            var doc = MarkdownReader.Parse(markdown);
            var html = doc.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                CodeBlockHtmlRenderer = (block, _) => $"<aside class=\"code-proxy\">{System.Net.WebUtility.HtmlEncode(block.Content)}</aside>"
            });

            Assert.Contains("<ul>", html, StringComparison.Ordinal);
            Assert.Contains("<aside class=\"code-proxy\">hello from list</aside>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<pre><code class=\"language-text\">", html, StringComparison.Ordinal);
        }
    }
}
