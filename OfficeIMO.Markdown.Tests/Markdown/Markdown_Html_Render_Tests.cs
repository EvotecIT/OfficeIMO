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
        public void Default_VisualTheme_Emits_WordLike_Html_Css() {
            string html = MarkdownDoc.Create()
                .H1("Title")
                .P("Hello")
                .ToHtmlDocument();

            Assert.Contains("article.markdown-body { color: #1f2937; background: #ffffff; }", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("article.markdown-body h1", html, StringComparison.Ordinal);
            Assert.Contains("--md-heading: #111827", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("--md-accent: #2563eb", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Default_VisualTheme_Can_Be_Disabled_For_Plain_Html_Css() {
            string html = MarkdownDoc.Create()
                .H1("Title")
                .P("Hello")
                .ToHtmlDocument(new HtmlOptions {
                    ApplyDefaultTheme = false
                });

            Assert.DoesNotContain("article.markdown-body { color: #1f2937; background: #ffffff; }", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("--md-heading: #111827", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void External_Css_Writes_Sidecar() {
            var temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            var path = Path.Combine(temp, "sample.html");
            var doc = MarkdownDoc.Create().H1("Title");
            doc.SaveAsHtml(path, new HtmlOptions { Title = "Doc", Style = HtmlStyle.Clean, CssDelivery = CssDelivery.ExternalFile });
            Assert.True(File.Exists(path));
            var cssPath = Path.ChangeExtension(path, ".css");
            Assert.True(File.Exists(cssPath));
        }

        [Fact]
        public void Html_Rendering_Does_Not_Mutate_Reusable_Options() {
            var options = new HtmlOptions {
                Kind = HtmlKind.Document,
                Theme = MarkdownVisualTheme.Report(),
                Prism = new PrismOptions { Enabled = true }
            };
            options.AdditionalCssHrefs.Add("https://example.com/site.css");

            string fragment = MarkdownDoc.Create().H1("Title").ToHtmlFragment(options);

            Assert.DoesNotContain("<html", fragment, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(HtmlKind.Document, options.Kind);
            Assert.Single(options.AdditionalCssHrefs);
            Assert.NotNull(options.Theme);
            Assert.True(options.Prism.Enabled);
        }

        [Fact]
        public async System.Threading.Tasks.Task External_Css_Async_Saves_Are_Operation_Scoped_And_Cancellable() {
            string temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            var options = new HtmlOptions {
                Kind = HtmlKind.Document,
                Style = HtmlStyle.Clean,
                CssDelivery = CssDelivery.ExternalFile
            };
            MarkdownDoc doc = MarkdownDoc.Create().H1("Title");
            string first = Path.Combine(temp, "first.html");
            string second = Path.Combine(temp, "second.html");

            await System.Threading.Tasks.Task.WhenAll(
                doc.SaveAsHtmlAsync(first, options),
                doc.SaveAsHtmlAsync(second, options));

            Assert.True(File.Exists(first));
            Assert.True(File.Exists(Path.ChangeExtension(first, ".css")));
            Assert.True(File.Exists(second));
            Assert.True(File.Exists(Path.ChangeExtension(second, ".css")));
            Assert.Equal(HtmlKind.Document, options.Kind);

            string cancelled = Path.Combine(temp, "cancelled.html");
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                doc.SaveAsHtmlAsync(cancelled, options, new System.Threading.CancellationToken(canceled: true)));
            Assert.False(File.Exists(cancelled));
        }

        [Fact]
        public void Shared_VisualTheme_Emits_Consistent_Html_Css() {
            MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
                .WithColorScheme(MarkdownColorSchemeKind.Emerald)
                .WithColors(accent: "SeaGreen", heading: "#064e3b", text: "#102030", background: "#f7fbff");
            theme.Table.BorderWidth = 1.2;
            theme.Table.CellPaddingX = 11;
            theme.Table.EmphasizeHeader = false;
            theme.Table.UseRowStripes = false;

            string html = MarkdownDoc.Create()
                .H1("Theme")
                .Table(t => t.Headers("Name", "Value").Row("Accent", "Emerald"))
                .ToHtmlDocument(new HtmlOptions {
                    Title = "Shared Theme",
                    Theme = theme,
                    Kind = HtmlKind.Document
                });

            Assert.Contains("body { --md-heading: #064e3b", html, StringComparison.Ordinal);
            Assert.Contains("article.markdown-body h1", html, StringComparison.Ordinal);
            Assert.Contains("--md-heading: #064e3b", html, StringComparison.Ordinal);
            Assert.Contains("article.markdown-body { color: #102030; background: #f7fbff; }", html, StringComparison.Ordinal);
            Assert.Contains("border-color: #a7f3d0", html, StringComparison.Ordinal);
            Assert.Contains("border-width: 1.2px", html, StringComparison.Ordinal);
            Assert.Contains("article.markdown-body th { background: transparent; color: inherit; }", html, StringComparison.Ordinal);
            Assert.Contains("article.markdown-body th, article.markdown-body td { border-color: #a7f3d0; border-width: 1.2px; padding: 5px 11px; }", html, StringComparison.Ordinal);
            Assert.Contains("padding: 5px 11px", html, StringComparison.Ordinal);
            Assert.Contains("tbody tr:nth-child(2n) { background-color: transparent; }", html, StringComparison.Ordinal);
            Assert.DoesNotContain("tbody tr:nth-child(even)", html, StringComparison.Ordinal);
        }

        [Fact]
        public void ColorOverridesApplyToHtmlThemeVariables() {
            string html = MarkdownDoc.Create()
                .H1("Legacy colors")
                .ToHtmlDocument(new HtmlOptions {
                    Title = "Custom colors",
                    ColorOverrides = new MarkdownHtmlColorOverrides {
                        HeadingLight = "SeaGreen",
                        AccentLight = "#123456"
                    },
                    Kind = HtmlKind.Document
                });

            Assert.Contains("body { --md-heading: #2e8b57; --md-accent: #123456;", html, StringComparison.Ordinal);
            Assert.Contains("article.markdown-body h1", html, StringComparison.Ordinal);
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
        public void HtmlOptions_Can_Disable_Automatic_Heading_Identifiers() {
            var doc = OfficeIMO.Markdown.MarkdownReader.Parse("# Alpha");

            string html = doc.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                AutoHeadingIdentifiers = false,
                IncludeAnchorLinks = true
            });

            Assert.Equal("<h1>Alpha</h1>", html);
        }

        [Fact]
        public void HtmlOptions_Supports_Markdig_Default_Heading_Identifier_Style() {
            const string markdown = """
# Привет мир

# Привет мир

# a_b c.d
""";

            string html = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                HeadingIdentifierStyle = MarkdownHeadingIdentifierStyle.MarkdigDefault
            });

            Assert.Contains("<h1 id=\"section\">Привет мир</h1>", html, StringComparison.Ordinal);
            Assert.Contains("<h1 id=\"section-1\">Привет мир</h1>", html, StringComparison.Ordinal);
            Assert.Contains("<h1 id=\"a_b-c.d\">a_b c.d</h1>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void HtmlOptions_Applies_Heading_Identifier_Style_To_Nested_Headings() {
            const string markdown = "- # a_b c.d";

            string html = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                HeadingIdentifierStyle = MarkdownHeadingIdentifierStyle.MarkdigDefault
            });

            Assert.Contains("<h1 id=\"a_b-c.d\">a_b c.d</h1>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("id=\"a-b-c-d\"", html, StringComparison.Ordinal);
        }

        [Fact]
        public void HtmlOptions_Can_Render_NonAscii_Text_Literally_For_Markdig_Style_Output() {
            const string markdown = """
åbc 1 < 2 & 'quote'

[ålink](https://example.com)

```
åcode < &
```
""";

            var defaultHtml = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null
            });
            var markdigStyleHtml = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            });

            Assert.Contains("&#229;bc", defaultHtml, StringComparison.Ordinal);
            Assert.Contains("&#229;link", defaultHtml, StringComparison.Ordinal);
            Assert.Contains("&#229;code", defaultHtml, StringComparison.Ordinal);
            Assert.Contains("åbc 1 &lt; 2 &amp; &#39;quote&#39;", markdigStyleHtml, StringComparison.Ordinal);
            Assert.Contains("<a href=\"https://example.com\">ålink</a>", markdigStyleHtml, StringComparison.Ordinal);
            Assert.Contains("åcode &lt; &amp;\n", markdigStyleHtml, StringComparison.Ordinal);
        }

        [Fact]
        public void HtmlOptions_Can_Render_NonAscii_Helper_Text_Literally_For_Markdig_Style_Output() {
            var toc = new TocBlock {
                NormalizeLevels = true
            };
            toc.Entries.Add(new TocBlock.Entry {
                Level = 2,
                Text = "åHeading",
                Anchor = "åheading"
            });

            var doc = MarkdownDoc.Create()
                .Add(toc)
                .H2("åHeading")
                .Add(new ImageBlock("https://example.com/image.png", "åAlt", "åTitle") {
                    Caption = "åCaption"
                });

            string html = doc.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false,
                IncludeAnchorLinks = true,
                AnchorIcon = "åAnchor",
                BackToTopLinks = true,
                BackToTopMinLevel = 2,
                BackToTopText = "åTop"
            });

            Assert.Contains(">åHeading</a>", html, StringComparison.Ordinal);
            Assert.Contains(">åAnchor</a>", html, StringComparison.Ordinal);
            Assert.Contains(">åTop</a>", html, StringComparison.Ordinal);
            Assert.Contains("<div class=\"caption\">åCaption</div>", html, StringComparison.Ordinal);
            Assert.Contains("alt=\"åAlt\"", html, StringComparison.Ordinal);
            Assert.Contains("title=\"åTitle\"", html, StringComparison.Ordinal);

            string documentHtml = doc.ToHtmlDocument(new HtmlOptions {
                Title = "åDocument",
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            });
            Assert.Contains("<title>åDocument</title>", documentHtml, StringComparison.Ordinal);

            string blockedLinkedImage = OfficeIMO.Markdown.MarkdownReader.Parse("[![åBadge](https://img.example/badge.svg)](https://example.com)")
                .ToHtmlFragment(new HtmlOptions {
                    Style = HtmlStyle.Plain,
                    CssDelivery = CssDelivery.None,
                    BodyClass = null,
                    EscapeNonAsciiText = false,
                    BlockExternalHttpImages = true
                });
            Assert.Contains(">åBadge</a>", blockedLinkedImage, StringComparison.Ordinal);
        }

        [Fact]
        public void HtmlOptions_Can_Render_NonAscii_Attribute_Text_Literally_For_Markdig_Style_Output() {
            var paragraph = new ParagraphBlock(new InlineSequence()
                .Link("ålink", "https://example.com/docs", "åLink title")
                .Image("åImage alt", "https://example.com/image.png", "åImage title")
                .ImageLink("åBadge", "https://img.example/badge.svg", "https://example.com", "åBadge title", "åBadge link"));
            var image = new ImageBlock("https://example.com/photo.png", "åBlock alt", "åBlock title", linkUrl: "https://example.com/photo", linkTitle: "åBlock link");
            image.PictureSources.Add(new ImagePictureSource(
                "https://example.com/photo.webp",
                "(min-width: 960px)",
                "image/webp",
                "åwide",
                "https://example.com/photo.webp 1x"));
            var doc = MarkdownDoc.Create().Add(paragraph).Add(image);

            string html = doc.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            });

            Assert.Contains("title=\"åLink title\"", html, StringComparison.Ordinal);
            Assert.Contains("alt=\"åImage alt\" title=\"åImage title\"", html, StringComparison.Ordinal);
            Assert.Contains("title=\"åBadge link\"", html, StringComparison.Ordinal);
            Assert.Contains("alt=\"åBadge\" title=\"åBadge title\"", html, StringComparison.Ordinal);
            Assert.Contains("title=\"åBlock link\"", html, StringComparison.Ordinal);
            Assert.Contains("alt=\"åBlock alt\" title=\"åBlock title\"", html, StringComparison.Ordinal);
            Assert.Contains("sizes=\"åwide\"", html, StringComparison.Ordinal);
            Assert.DoesNotContain("&#229;", html, StringComparison.Ordinal);
        }

        [Fact]
        public void HtmlOptions_Can_Render_NonAscii_Generated_Attribute_Metadata_Literally_For_Markdig_Style_Output() {
            var toc = new TocBlock();
            toc.Entries.Add(new TocBlock.Entry {
                Level = 2,
                Text = "åToc",
                Anchor = "åtoc"
            });

            var doc = MarkdownDoc.Create()
                .H2("åHeading")
                .Add(toc)
                .Add(new ParagraphBlock(new InlineSequence()
                    .Link("åExternal", "https://example.com/docs", linkTarget: "åtarget", linkRel: "årel")
                    .Link("åPlainExternal", "https://example.com/plain")
                    .Image("åImage", "https://example.com/image.png")
                    .FootnoteRef("ånote")))
                .Add(new CodeBlock("ålang", "åcode"))
                .Add(new SemanticFencedBlock("chart", "åsemantic", "åpayload"))
                .Add(new CalloutBlock("ånote", "åTitle", "åBody"))
                .Add(new FootnoteDefinitionBlock("ånote", "åFootnote"));

            var options = new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false,
                HeadingIdentifierStyle = MarkdownHeadingIdentifierStyle.GitHub,
                IncludeAnchorLinks = true,
                ExternalLinksReferrerPolicy = "åpolicy",
                ImagesReferrerPolicy = "åimage-policy"
            };

            string html = doc.ToHtmlFragment(options);

            Assert.Contains("id=\"åheading\"", html, StringComparison.Ordinal);
            Assert.Contains("href=\"#åheading\"", html, StringComparison.Ordinal);
            Assert.Contains("data-anchor-id=\"åheading\"", html, StringComparison.Ordinal);
            Assert.Contains("href=\"#åtoc\">åToc</a>", html, StringComparison.Ordinal);
            Assert.Contains("target=\"åtarget\"", html, StringComparison.Ordinal);
            Assert.Contains("rel=\"årel\"", html, StringComparison.Ordinal);
            Assert.Contains("referrerpolicy=\"åpolicy\"", html, StringComparison.Ordinal);
            Assert.Contains("referrerpolicy=\"åimage-policy\"", html, StringComparison.Ordinal);
            Assert.Contains("id=\"fnref:ånote\"", html, StringComparison.Ordinal);
            Assert.Contains("id=\"fn:ånote\"", html, StringComparison.Ordinal);
            Assert.Contains("class=\"language-ålang\"", html, StringComparison.Ordinal);
            Assert.Contains("class=\"language-åsemantic\"", html, StringComparison.Ordinal);
            Assert.Contains("class=\"callout ånote\"", html, StringComparison.Ordinal);
            Assert.DoesNotContain("&#229;", html, StringComparison.Ordinal);

            string documentHtml = doc.ToHtmlDocument(new HtmlOptions {
                Title = "åDocument",
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.LinkHref,
                CssHref = "https://cdn.example.com/å.css",
                BodyClass = "å-body",
                EscapeNonAsciiText = false,
                AssetMode = AssetMode.Online,
                AdditionalCssHrefs = { "https://cdn.example.com/extra å.css" },
                AdditionalJsHrefs = { "https://cdn.example.com/å.js" }
            });

            Assert.Contains("<title>åDocument</title>", documentHtml, StringComparison.Ordinal);
            Assert.Contains("<article class=\"å-body\">", documentHtml, StringComparison.Ordinal);
            Assert.Contains("href=\"https://cdn.example.com/%C3%A5.css\"", documentHtml, StringComparison.Ordinal);
            Assert.Contains("href=\"https://cdn.example.com/extra%20%C3%A5.css\"", documentHtml, StringComparison.Ordinal);
            Assert.Contains("src=\"https://cdn.example.com/%C3%A5.js\"", documentHtml, StringComparison.Ordinal);
            Assert.DoesNotContain("&#229;", documentHtml, StringComparison.Ordinal);

            var merged = HtmlAssetMerger.Build(
                new[] {
                    new[] {
                        new HtmlAsset("åcss", HtmlAssetKind.Css, "https://cdn.example.com/å.css", inline: null) {
                            Media = "screen and (min-width: åpx)"
                        },
                        new HtmlAsset("åjs", HtmlAssetKind.Js, "https://cdn.example.com/å.js", inline: null)
                    }
                },
                new HtmlOptions {
                    EscapeNonAsciiText = false
                });

            Assert.Contains("data-asset-id=\"åcss\"", merged.headLinks, StringComparison.Ordinal);
            Assert.Contains("media=\"screen and (min-width: åpx)\"", merged.headLinks, StringComparison.Ordinal);
            Assert.Contains("href=\"https://cdn.example.com/%C3%A5.css\"", merged.headLinks, StringComparison.Ordinal);
            Assert.Contains("data-asset-id=\"åjs\"", merged.headLinks, StringComparison.Ordinal);
            Assert.Contains("src=\"https://cdn.example.com/%C3%A5.js\"", merged.headLinks, StringComparison.Ordinal);
            Assert.DoesNotContain("&#229;", merged.headLinks, StringComparison.Ordinal);
        }

        [Fact]
        public void Link_Html_Preserves_Ipv6_Authority_Brackets() {
            const string markdown = "[loopback](http://[::1]/)";

            string html = OfficeIMO.Markdown.MarkdownReader.Parse(markdown)
                .ToHtmlFragment(new HtmlOptions {
                    Style = HtmlStyle.Plain,
                    CssDelivery = CssDelivery.None,
                    BodyClass = null
                });

            Assert.Contains("href=\"http://[::1]/\"", html, StringComparison.Ordinal);
            Assert.DoesNotContain("href=\"http://%5B::1%5D/\"", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void GitHubFlavoredMarkdown_Html_Profile_Uses_GitHub_Heading_Identifiers() {
            const string markdown = """
# Hello World!

# Hello World!

# Привет мир

# a_b c.d
""";

            var options = HtmlOptions.CreateGitHubFlavoredMarkdownProfile();
            string html = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile())
                .ToHtmlFragment(options);

            Assert.True(options.AutoHeadingIdentifiers);
            Assert.Equal(MarkdownHeadingIdentifierStyle.GitHub, options.HeadingIdentifierStyle);
            Assert.Contains("<h1 id=\"hello-world\">Hello World!</h1>", html, StringComparison.Ordinal);
            Assert.Contains("<h1 id=\"hello-world-1\">Hello World!</h1>", html, StringComparison.Ordinal);
            Assert.Contains("<h1 id=\"привет-мир\">Привет мир</h1>", html, StringComparison.Ordinal);
            Assert.Contains("<h1 id=\"a_b-cd\">a_b c.d</h1>", html, StringComparison.Ordinal);
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
            var doc = OfficeIMO.Markdown.MarkdownReader.Parse(markdown);
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
