using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        private static string AssetImagePath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));

        [Fact]
        public void MarkdownToWord_CodeBlock_EveryLineUsesTheMonospaceFont() {
            var markdown = Omd.MarkdownDocument.Parse("```bash\nlumen init my-notes\n\n    cd my-notes\n```\n");

            using var document = markdown.ToWord();
            WordParagraph code = document.Paragraphs.First(paragraph => paragraph.StyleId?.StartsWith("CodeLang_", StringComparison.Ordinal) == true);
            Assert.Single(document.Paragraphs.Where(paragraph => paragraph.StyleId?.StartsWith("CodeLang_", StringComparison.Ordinal) == true)
                .Select(paragraph => paragraph._paragraph).Distinct());
            Paragraph element = code._paragraph;
            Run[] runs = element.Elements<Run>().ToArray();

            Assert.Equal("CodeLang_bash", code.StyleId);
            Assert.All(runs, run => Assert.False(string.IsNullOrEmpty(run.RunProperties?.RunFonts?.Ascii?.Value)));
            Assert.Single(runs.Select(run => run.RunProperties!.RunFonts!.Ascii!.Value).Distinct());
            // Blank lines and indentation survive as line breaks and preserved spaces.
            Assert.Equal(2, element.Descendants<Break>().Count());
            Assert.Contains(element.Descendants<Text>(), text => text.Text == "    cd my-notes");
            Assert.True(document.StyleExists("CodeLang_bash"));
        }

        [Theory]
        [InlineData("- Setup\n\n  ```csharp\n  run();\n  ```\n", 720)]
        [InlineData("- Outer\n  - Inner\n\n    ```csharp\n    run();\n    ```\n", 1440)]
        [InlineData("> - Setup\n>\n>   ```csharp\n>   run();\n>   ```\n", 1440)]
        public void MarkdownToWord_CodeBlocksRetainListAndQuoteIndentation(string markdown, int expectedTwips) {
            using WordDocument source = Omd.MarkdownDocument.Parse(markdown).ToWord();
            using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
            Paragraph code = Assert.Single(document.Paragraphs
                .Where(paragraph => paragraph.StyleId == "CodeLang_csharp")
                .Select(paragraph => paragraph._paragraph).Distinct());

            Assert.Equal(expectedTwips.ToString(System.Globalization.CultureInfo.InvariantCulture),
                code.ParagraphProperties!.Indentation!.Left!.Value);
            Assert.Equal("run();", string.Concat(code.Descendants<Text>().Select(text => text.Text)));
            Assert.Empty(document.ValidateDocument());
        }

        [Fact]
        public void MarkdownToWordToMarkdown_KeepsFencedCode() {
            var markdown = Omd.MarkdownDocument.Parse("Intro\n\n```json\n{\n  \"a\": 1\n}\n```\n");

            using var document = markdown.ToWord();
            string roundTrip = document.ToMarkdown().ToMarkdown().Replace("\r\n", "\n");

            Assert.Contains("```json\n{\n  \"a\": 1\n}\n```", roundTrip);
        }

        [Fact]
        public void MarkdownToWordToMarkdown_KeepsHorizontalRulesAndInternalLinks() {
            var markdown = Omd.MarkdownDocument.Parse("# Lumen\n\nSee [quick start](#quick-start).\n\n## Quick start\n\nText.\n\n---\n\nLicensed under MIT.\n");

            using var document = markdown.ToWord();
            string roundTrip = document.ToMarkdown().ToMarkdown().Replace("\r\n", "\n");

            Assert.Contains("[quick start](#quick-start)", roundTrip);
            Assert.Contains("\n---\n", roundTrip);
            Assert.Contains("Licensed under MIT.", roundTrip);
        }

        [Fact]
        public void MarkdownToWord_InternalLinksTargetHeadingBookmarks() {
            var markdown = Omd.MarkdownDocument.Parse("# Lumen\n\nSee [the changelog](#changelog) and [quick start](#quick-start).\n\n## Quick start\n\nText.\n\n## Changelog\n\nText.\n");

            using var document = markdown.ToWord(new WordMarkdownImportOptions(), out var report);
            var body = document._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
            string[] bookmarks = body.Descendants<BookmarkStart>().Select(bookmark => bookmark.Name!.Value!).ToArray();
            Hyperlink[] links = body.Descendants<Hyperlink>().ToArray();

            Assert.Contains("changelog", bookmarks);
            Assert.Contains("quick_start", bookmarks);
            Assert.Equal(new[] { "changelog", "quick_start" }, links.Select(link => link.Anchor?.Value).ToArray());
            Assert.All(links, link => Assert.Null(link.Id));
            Assert.Empty(document._wordprocessingDocument.MainDocumentPart!.HyperlinkRelationships);
            // The bookmark wraps the heading paragraph content.
            Paragraph heading = body.Descendants<BookmarkStart>().First(bookmark => bookmark.Name == "changelog").Parent as Paragraph
                ?? throw new InvalidOperationException("Bookmark is not inside a paragraph.");
            Assert.Equal("Changelog", string.Concat(heading.Descendants<Text>().Select(text => text.Text)));
            Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Message.Contains("Internal link"));
        }

        [Fact]
        public void MarkdownToWord_RelativeImageResolvesAgainstBaseUriAndSourceFile() {
            string directory = Path.Combine(Path.GetTempPath(), "officeimo-md-images-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(directory, "assets"));
            try {
                File.Copy(AssetImagePath, Path.Combine(directory, "assets", "pipeline.png"));
                string markdownPath = Path.Combine(directory, "readme.md");
                File.WriteAllText(markdownPath, "# Doc\n\n![Pipeline](assets/pipeline.png)\n");

                foreach (string? baseUri in new[] { directory, new Uri(directory + Path.DirectorySeparatorChar).AbsoluteUri, null }) {
                    using var document = Omd.MarkdownDocument.Load(markdownPath).ToWord(
                        new WordMarkdownImportOptions { BaseUri = baseUri, AllowLocalImages = true }, out var report);
                    Assert.Single(document.Images);
                    Assert.False(report.HasLoss);
                }
            } finally {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void MarkdownToWord_ImageThatCannotBeEmbeddedReportsLoss() {
            string directory = Path.Combine(Path.GetTempPath(), "officeimo-md-images-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try {
                File.Copy(AssetImagePath, Path.Combine(directory, "pipeline.png"));

                using var disabled = Omd.MarkdownDocument.Parse("![Pipeline](pipeline.png)").ToWord(
                    new WordMarkdownImportOptions { BaseUri = directory }, out var disabledReport);
                Assert.Empty(disabled.Images);
                Assert.True(disabledReport.HasLoss);
                Assert.Contains(disabledReport.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ImageSkippedByPolicy
                    && diagnostic.LossKind == OfficeConversionLossKind.Omission);

                using var missing = Omd.MarkdownDocument.Parse("![Missing](missing.png)").ToWord(
                    new WordMarkdownImportOptions { BaseUri = directory, AllowLocalImages = true }, out var missingReport);
                Assert.Empty(missing.Images);
                Assert.Contains(missingReport.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ImageLoadFailed);
            } finally {
                Directory.Delete(directory, recursive: true);
            }
        }

        [Fact]
        public void MarkdownToWord_HeadingStylesHaveAThemeHierarchy() {
            using var document = Omd.MarkdownDocument.Parse("# One\n\n## Two\n\n### Three\n\nBody").ToWord();
            var styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            OfficeTheme theme = OfficeTheme.Default;
            double previous = double.MaxValue;
            foreach ((string id, OfficeLength size) in new[] {
                ("Heading1", theme.TypeScale.Heading1), ("Heading2", theme.TypeScale.Heading2), ("Heading3", theme.TypeScale.Heading3) }) {
                Style style = styles.Elements<Style>().Single(item => item.StyleId == id);
                StyleRunProperties run = style.StyleRunProperties!;
                double points = double.Parse(run.FontSize!.Val!.Value!, System.Globalization.CultureInfo.InvariantCulture) / 2D;
                Assert.Equal(size.ToPoints(), points, 1);
                Assert.True(points < previous);
                previous = points;
                Assert.NotNull(run.Bold);
                Assert.Equal(theme.Fonts.Heading.Family, run.RunFonts!.Ascii!.Value);
                Assert.NotNull(style.StyleParagraphProperties!.KeepNext);
            }
            Assert.True(previous > theme.TypeScale.Body.ToPoints());
        }

        [Fact]
        public void MarkdownToWord_NormalStyleAndNumericColumnsFollowTheTheme() {
            using var document = Omd.MarkdownDocument.Parse("| Region | Revenue |\n|---|---|\n| EMEA | 1,250 |\n| APAC | 980 |\n").ToWord();
            var styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style normal = styles.Elements<Style>().Single(item => item.StyleId == "Normal");
            Assert.Equal(OfficeTheme.Default.Fonts.Body.Family, normal.StyleRunProperties!.RunFonts!.Ascii!.Value);
            Assert.Equal((OfficeTheme.Default.TypeScale.Body.ToPoints() * 2D).ToString(System.Globalization.CultureInfo.InvariantCulture),
                normal.StyleRunProperties.FontSize!.Val!.Value);

            WordTable table = Assert.Single(document.Tables);
            Assert.Equal(WordParagraphAlignment.Right, table.Rows[1].Cells[1].Paragraphs.First(p => !string.IsNullOrEmpty(p.Text)).ParagraphAlignment);
            Assert.NotEqual(WordParagraphAlignment.Right, table.Rows[1].Cells[0].Paragraphs.First(p => !string.IsNullOrEmpty(p.Text)).ParagraphAlignment);
        }

        [Fact]
        public void MarkdownToWord_TableColumnsFollowTheirContent() {
            var markdown = Omd.MarkdownDocument.Parse(
                "| Setting | Type | Description |\n|---|---|---|\n" +
                "| `title` | string | Site title shown in the header and browser tab of every generated page. |\n" +
                "| `output` | string | Folder that receives the generated site. Relative to the project root. |\n");

            using var document = markdown.ToWord();
            WordTable table = Assert.Single(document.Tables);
            var widths = table.Rows[0].Cells.Select(cell => {
                WordTableWidth width = Assert.IsType<WordTableWidth>(cell.Width);
                Assert.Equal(WordTableWidthKind.Percent, width.Kind);
                return width.PercentValue!.Value;
            }).ToArray();

            Assert.True(widths[2] > widths[0] * 2, string.Join(",", widths));
            Assert.True(widths[2] > widths[1] * 2, string.Join(",", widths));
        }
    }
}
