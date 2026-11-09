using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        private const string QualityTableMarkdown =
            "| Setting | Type | Default | Description |\n|:--|:--:|--:|:--|\n" +
            "| `title` | string | `\"Lumen site\"` | Site title shown in the header and browser tab. |\n" +
            "| `search.maxResults` | integer | `20` | Maximum number of hits returned per query. |\n";

        private static IEnumerable<HtmlRenderVisual> FlattenScene(IEnumerable<HtmlRenderVisual> visuals) {
            foreach (HtmlRenderVisual visual in visuals) {
                yield return visual;
                IEnumerable<HtmlRenderVisual>? children = visual switch {
                    HtmlRenderSemanticGroup semantic => semantic.Visuals,
                    HtmlRenderLogicalTextGroup logical => logical.Visuals,
                    HtmlRenderClipGroup clip => clip.Visuals,
                    HtmlRenderPathClipGroup pathClip => pathClip.Visuals,
                    HtmlRenderEffectGroup effect => effect.Visuals,
                    HtmlRenderFormField form => form.Visuals,
                    _ => null
                };
                if (children == null) continue;
                foreach (HtmlRenderVisual child in FlattenScene(children)) yield return child;
            }
        }

        private static HtmlRenderText[] RenderTexts(string html) =>
            HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html)).Pages
                .SelectMany(page => FlattenScene(page.Scene)).OfType<HtmlRenderText>().ToArray();

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void MarkdownHtml_SaveAsHtmlWritesACompleteDocumentWithOptionalPresentation(bool applyStyle) {
            string path = Path.Combine(Path.GetTempPath(), "officeimo-md-" + Guid.NewGuid().ToString("N") + ".html");
            try {
                var document = Omd.MarkdownDocument.Parse("# Lumen CLI\n\nBody.");
                if (applyStyle) {
                    document.SaveAsHtml(path, new Omd.MarkdownHtmlExportOptions {
                        Kind = Omd.MarkdownHtmlKind.Document,
                        Style = Omd.MarkdownHtmlStyle.Clean,
                        Title = "Lumen CLI"
                    });
                } else {
                    document.SaveAsHtml(path);
                }
                string html = File.ReadAllText(path);
                Assert.StartsWith("<!DOCTYPE html>", html.TrimStart(), StringComparison.OrdinalIgnoreCase);
                Assert.Contains("<meta charset", html, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("<title>Lumen CLI</title>", html);
                var parsed = new AngleSharp.Html.Parser.HtmlParser().ParseDocument(html);
                Assert.Equal("Lumen CLI", parsed.QuerySelector("h1")?.TextContent);
                Assert.Equal("Body.", parsed.QuerySelector("p")?.TextContent);
                if (applyStyle) {
                    Assert.Contains("<style", html, StringComparison.OrdinalIgnoreCase);
                } else {
                    Assert.DoesNotContain("<style", html, StringComparison.OrdinalIgnoreCase);
                }
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void MarkdownHtml_CleanStyleDiffersFromWordStyle() {
            var markdown = Omd.MarkdownDocument.Parse("# Title\n\nBody.");
            string clean = markdown.ToHtml(new Omd.MarkdownHtmlExportOptions { Kind = Omd.MarkdownHtmlKind.Document, Style = Omd.MarkdownHtmlStyle.Clean });
            string word = markdown.ToHtml(new Omd.MarkdownHtmlExportOptions { Kind = Omd.MarkdownHtmlKind.Document, Style = Omd.MarkdownHtmlStyle.Word });
            Assert.NotEqual(clean, word);
        }

        [Theory]
        [InlineData(Omd.MarkdownHtmlStyle.Clean)]
        [InlineData(Omd.MarkdownHtmlStyle.GithubLight)]
        [InlineData(Omd.MarkdownHtmlStyle.Word)]
        public void MarkdownHtml_TableCellsDoNotBreakShortWords(Omd.MarkdownHtmlStyle style) {
            string html = Omd.MarkdownDocument.Parse(QualityTableMarkdown).ToHtml(
                new Omd.MarkdownHtmlExportOptions { Kind = Omd.MarkdownHtmlKind.Document, Style = style });
            HtmlRenderText[] texts = RenderTexts(html);
            foreach (string word in new[] { "Type", "string", "integer", "search.maxResults" }) {
                Assert.Contains(texts, text => text.Text.Trim() == word);
            }
        }

        [Theory]
        [InlineData(Omd.MarkdownHtmlStyle.Clean)]
        [InlineData(Omd.MarkdownHtmlStyle.GithubLight)]
        [InlineData(Omd.MarkdownHtmlStyle.Word)]
        public void MarkdownHtml_CodeUsesAMonospaceStackAndBlocksHaveOneBackground(Omd.MarkdownHtmlStyle style) {
            string html = Omd.MarkdownDocument.Parse("Use `lumen`.\n\n```bash\nlumen build\nlumen serve\n```\n").ToHtml(
                new Omd.MarkdownHtmlExportOptions { Kind = Omd.MarkdownHtmlKind.Document, Style = style });
            HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html));
            HtmlRenderVisual[] scene = rendered.Pages.SelectMany(page => FlattenScene(page.Scene)).ToArray();
            HtmlRenderText inline = scene.OfType<HtmlRenderText>().Single(text => text.Text == "lumen");
            HtmlRenderText body = scene.OfType<HtmlRenderText>().First(text => text.Text.StartsWith("Use", StringComparison.Ordinal));
            Assert.NotEqual(body.Font.FamilyName, inline.Font.FamilyName);
            HtmlRenderText block = scene.OfType<HtmlRenderText>().First(text => text.Text.Contains("build"));
            Assert.NotEqual(body.Font.FamilyName, block.Font.FamilyName);
            Assert.Contains("monospace", html);
            Assert.Contains("pre code", html);
        }

        [Fact]
        public void HtmlThemeTokensRetainAuthoredSurfacesAndLinkColor() {
            var color = global::OfficeIMO.OfficeColor.Parse("#123456");
            OfficeTheme theme = OfficeTheme.Modern.WithSurfaces(OfficeTheme.Modern.Surfaces.With(
                surface: color, rule: color, accent1Tint: color, successFill: color, warningFill: color, dangerFill: color))
                .WithTextStyle(OfficeTheme.Modern.TextStyle.With(linkColor: color));

            using var source = Omd.MarkdownDocument.Parse("# Layered heading\n\n## Low-specificity heading\n\n" +
                "[Surface link](https://example.test/theme)\n\n```text\nSurface code\n```\n");
            string html = source.ToHtml(new Omd.MarkdownHtmlExportOptions {
                Kind = Omd.MarkdownHtmlKind.Document,
                Theme = Omd.MarkdownVisualTheme.FromOfficeTheme(theme),
                AssetMode = Omd.MarkdownAssetMode.Offline,
                AdditionalCssHrefs = { "https://example.test/author.css" },
                ExternalTextResolver = _ => "@layer report{h1{color:#008800}}:where(h2){color:#880000}"
            });

            foreach (string token in new[] { "surface", "rule", "accent-1-tint", "success-fill", "warning-fill", "danger-fill", "link" })
                Assert.Contains("--office-" + token + ": #123456;", html, StringComparison.Ordinal);
            HtmlRenderVisual[] scene = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html)).Pages
                .SelectMany(page => FlattenScene(page.Scene)).ToArray();
            Assert.Contains(scene.OfType<HtmlRenderShape>(), shape => shape.Shape.FillColor == color);
            Assert.Contains(scene.OfType<HtmlRenderText>(), text => text.LinkUri == "https://example.test/theme" && text.Color == color);
            Assert.Contains(scene.OfType<HtmlRenderText>(), text => text.Text == "Layered heading" && text.Color == OfficeColor.Parse("#008800"));
            Assert.Contains(scene.OfType<HtmlRenderText>(), text => text.Text == "Low-specificity heading" && text.Color == OfficeColor.Parse("#880000"));

            HtmlRenderPage RenderPage(string htmlDocument) {
                HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(htmlDocument), new HtmlRenderOptions {
                    Mode = HtmlRenderMode.Paged,
                    PageSize = new OfficePageSize(OfficeLength.Inches(390D / 96D), OfficeLength.Inches(780D / 96D)),
                    Margins = HtmlRenderMargins.All(0D)
                });
                return rendered.Pages[0];
            }
            OfficeColor CanvasPixel(string htmlDocument) {
                var image = OfficeDrawingRasterRenderer.Render(RenderPage(htmlDocument).CreateDrawing());
                return image.GetPixel(1, image.Height - 2);
            }
            using var darkSource = Omd.MarkdownDocument.Parse("Short dark document.");
            var darkOptions = new Omd.MarkdownHtmlExportOptions {
                Kind = Omd.MarkdownHtmlKind.Document,
                Theme = Omd.MarkdownVisualTheme.FromOfficeTheme(OfficeTheme.Dark)
            };
            Assert.Equal(OfficeColor.Parse("#0F1620"), CanvasPixel(darkSource.ToHtml(darkOptions)));
            darkOptions.AssetMode = Omd.MarkdownAssetMode.Offline;
            darkOptions.AdditionalCssHrefs.Add("https://example.test/canvas.css");
            darkOptions.ExternalTextResolver = _ => ":where(body){background-color:#665544}";
            Assert.Equal(OfficeColor.Parse("#665544"), CanvasPixel(darkSource.ToHtml(darkOptions)));
            string fragment = darkSource.ToHtml(new Omd.MarkdownHtmlExportOptions {
                Kind = Omd.MarkdownHtmlKind.Fragment,
                IncludeCssInFragment = true,
                Theme = Omd.MarkdownVisualTheme.FromOfficeTheme(OfficeTheme.Dark)
            });
            string host = "<!doctype html><html><body>Host content.</body></html>";
            string hosted = "<!doctype html><html><body>Host content." + fragment + "</body></html>";
            Assert.Equal(OfficeColor.White, CanvasPixel(hosted));
            HtmlRenderPage hostPage = RenderPage(host);
            HtmlRenderPage fragmentPage = RenderPage(hosted);
            Assert.Equal((hostPage.Width, hostPage.Height, hostPage.Margins.Top, hostPage.Margins.Right, hostPage.Margins.Bottom, hostPage.Margins.Left),
                (fragmentPage.Width, fragmentPage.Height, fragmentPage.Margins.Top, fragmentPage.Margins.Right, fragmentPage.Margins.Bottom, fragmentPage.Margins.Left));
        }

        [Fact]
        public void MarkdownHtml_ExplicitThemeUsesThemeTokensAndAlignsNumbers() {
            string html = Omd.MarkdownDocument.Parse("| Region | Revenue | Share |\n|---|---|---|\n| EMEA | 1,250 | 12.5% |\n| APAC | 980 | 9% |\n")
                .ToHtml(new Omd.MarkdownHtmlExportOptions {
                    Kind = Omd.MarkdownHtmlKind.Document,
                    Theme = Omd.MarkdownVisualTheme.FromOfficeTheme(OfficeTheme.Default)
                });
            OfficeTheme theme = OfficeTheme.Default;

            Assert.Contains("--office-accent-1: #" + theme.Palette.Accent1.ToRgbHex(), html);
            Assert.Contains("--office-font-body: \"" + theme.Fonts.Body.Family + "\"", html);
            Assert.Contains("tabular-nums", html);
            var parsed = new AngleSharp.Html.Parser.HtmlParser().ParseDocument(html);
            var firstRow = parsed.QuerySelectorAll("tbody tr").First().QuerySelectorAll("td").ToArray();
            Assert.DoesNotContain("text-align", firstRow[0].GetAttribute("style") ?? string.Empty);
            Assert.Contains("text-align:right", firstRow[1].GetAttribute("style") ?? string.Empty);
            Assert.Contains("text-align:right", firstRow[2].GetAttribute("style") ?? string.Empty);
            Assert.Contains("text-align:right", parsed.QuerySelectorAll("thead th").ElementAt(1).GetAttribute("style") ?? string.Empty);

            // The default plain output does not infer alignment from the cell values.
            string plain = Omd.MarkdownDocument.Parse("| A | B |\n|---|---|\n| x | 1 |\n")
                .ToHtml();
            Assert.DoesNotContain("text-align", plain);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        public void MarkdownHtml_TaskItemsKeepOrdinaryBulletsAndVisibleCheckboxes(bool authoredTaskMarker, bool applyStyle) {
            string markdown = "- **Fixed:** search\n- [x] Signed"
                + (authoredTaskMarker ? " {style=\"list-style:disc;color:#123456\"}" : string.Empty)
                + "\n- [ ] ARM64\n";
            var document = authoredTaskMarker
                ? Omd.MarkdownDocument.Parse(markdown, new Omd.MarkdownLoadOptions { GenericAttributes = true })
                : Omd.MarkdownDocument.Parse(markdown);
            // The CommonMark writer preserves attributed list items through its public raw-HTML fallback.
            string html = authoredTaskMarker
                ? document.ToMarkdown(Omd.MarkdownSaveOptions.CreateCommonMarkProfile())
                : document.ToHtml(new Omd.MarkdownHtmlExportOptions {
                    Kind = Omd.MarkdownHtmlKind.Document,
                    Style = applyStyle ? Omd.MarkdownHtmlStyle.Clean : Omd.MarkdownHtmlStyle.Plain
                });
            if (applyStyle) {
                Assert.Contains("<style", html, StringComparison.OrdinalIgnoreCase);
            } else {
                Assert.DoesNotContain("<style", html, StringComparison.OrdinalIgnoreCase);
            }
            var parsed = new AngleSharp.Html.Parser.HtmlParser().ParseDocument(html);
            var items = parsed.QuerySelectorAll("li");
            Assert.Equal(applyStyle ? null : authoredTaskMarker ? "list-style:none;list-style:disc;color:#123456" : "list-style:none;",
                items[1].GetAttribute("style"));
            HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html));
            HtmlRenderVisual[] scene = rendered.Pages.SelectMany(page => FlattenScene(page.Scene)).ToArray();

            Assert.Equal(authoredTaskMarker ? 2 : 1,
                scene.OfType<HtmlRenderText>().Count(text => text.SemanticRole == "list-marker"));
            HtmlRenderFormField[] boxes = scene.OfType<HtmlRenderFormField>().ToArray();
            Assert.Equal(2, boxes.Length);
            Assert.All(boxes, box => {
                Assert.True(box.Width >= 10D, box.Width.ToString(System.Globalization.CultureInfo.InvariantCulture));
                Assert.True(box.Height >= 10D, box.Height.ToString(System.Globalization.CultureInfo.InvariantCulture));
            });
        }
    }
}
