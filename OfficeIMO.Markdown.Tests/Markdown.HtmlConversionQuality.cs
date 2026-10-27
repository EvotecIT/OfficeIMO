using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        [Fact]
        public void MarkdownHtml_SaveAsHtmlWithoutOptionsWritesACompleteDocument() {
            string path = Path.Combine(Path.GetTempPath(), "officeimo-md-" + Guid.NewGuid().ToString("N") + ".html");
            try {
                Omd.MarkdownDocument.Parse("# Lumen CLI\n\nBody.").SaveAsHtml(path);
                string html = File.ReadAllText(path);
                Assert.StartsWith("<!DOCTYPE html>", html.TrimStart(), StringComparison.OrdinalIgnoreCase);
                Assert.Contains("<meta charset", html, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("<style", html, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("<title>Lumen CLI</title>", html);
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

            string css = global::OfficeIMO.Internal.OfficeThemeCssTokens.Build(theme, ":root");

            foreach (string token in new[] { "surface", "rule", "accent-1-tint", "success-fill", "warning-fill", "danger-fill", "link" })
                Assert.Contains("--office-" + token + ": #123456;", css, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownHtml_DefaultStyleUsesThemeTokensAndAlignsNumbers() {
            string html = Omd.MarkdownDocument.Parse("| Region | Revenue | Share |\n|---|---|---|\n| EMEA | 1,250 | 12.5% |\n| APAC | 980 | 9% |\n")
                .ToHtml(new Omd.MarkdownHtmlExportOptions { Kind = Omd.MarkdownHtmlKind.Document });
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

            // Source alignment stays exact for Plain output.
            string plain = Omd.MarkdownDocument.Parse("| A | B |\n|---|---|\n| x | 1 |\n")
                .ToHtml(new Omd.MarkdownHtmlExportOptions { Style = Omd.MarkdownHtmlStyle.Plain });
            Assert.DoesNotContain("text-align", plain);
        }

        [Fact]
        public void MarkdownHtml_TaskItemsKeepOrdinaryBulletsAndVisibleCheckboxes() {
            string html = Omd.MarkdownDocument.Parse("- **Fixed:** search\n- [x] Signed\n- [ ] ARM64\n").ToHtml(
                new Omd.MarkdownHtmlExportOptions { Kind = Omd.MarkdownHtmlKind.Document });
            HtmlRenderDocument rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html));
            HtmlRenderVisual[] scene = rendered.Pages.SelectMany(page => FlattenScene(page.Scene)).ToArray();

            Assert.Single(scene.OfType<HtmlRenderText>(), text => text.SemanticRole == "list-marker");
            HtmlRenderFormField[] boxes = scene.OfType<HtmlRenderFormField>().ToArray();
            Assert.Equal(2, boxes.Length);
            Assert.All(boxes, box => {
                Assert.True(box.Width >= 10D, box.Width.ToString(System.Globalization.CultureInfo.InvariantCulture));
                Assert.True(box.Height >= 10D, box.Height.ToString(System.Globalization.CultureInfo.InvariantCulture));
            });
        }
    }
}
