using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        public static IEnumerable<object[]> VisualMarkdownFixtures() {
            yield return new object[] { "mermaid-summary.md" };
            yield return new object[] { "vis-network-summary.md" };
            yield return new object[] { "chart-summary.md" };
        }

        [Theory]
        [MemberData(nameof(VisualMarkdownFixtures))]
        public void MarkdownToWord_VisualFixtures_AvoidEscapeArtifacts_AndFitImages(string fixtureName) {
            string fixturePath = Path.Combine(AppContext.BaseDirectory, "Documents", "MarkdownVisuals", fixtureName);
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string markdown = File.ReadAllText(fixturePath).Replace("{{LOCAL_IMAGE}}", imagePath);

            using var document = markdown.LoadFromMarkdown(new MarkdownToWordOptions {
                AllowLocalImages = true,
                PreferNarrativeSingleLineDefinitions = true,
                FitImagesToPageContentWidth = true,
                DefaultPageSize = WordPageSize.Letter,
                ImageLayout = {
                    AllowUpscale = true
                }
            });

            Assert.Single(document.Images);
            Assert.InRange(document.Images[0].Width ?? 0, 1, 625);

            var bodyText = string.Join("\n", document.Paragraphs.Select(p => p.Text));
            Assert.DoesNotContain("\\:", bodyText, StringComparison.Ordinal);
            Assert.DoesNotContain("\\n", bodyText, StringComparison.Ordinal);
            Assert.Contains("Interpretation:", bodyText, StringComparison.Ordinal);
        }
    }
}
