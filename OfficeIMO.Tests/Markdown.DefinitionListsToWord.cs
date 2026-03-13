using System;
using System.Linq;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_DefinitionListLine_RendersStrongDefinitionWithoutLiteralMarkers() {
            const string markdown = "Short answer: **no — nothing is failed** ✅";

            using var document = markdown.LoadFromMarkdown(new MarkdownToWordOptions());
            var paragraph = document.Paragraphs.First(p => p.Text.Contains("Short answer", StringComparison.Ordinal));
            var runs = paragraph.GetRuns();
            var combinedRunText = string.Concat(runs.Select(run => run.Text));

            Assert.DoesNotContain("**", combinedRunText, StringComparison.Ordinal);
            Assert.Contains("Short answer", combinedRunText, StringComparison.Ordinal);
            Assert.Contains("no — nothing is failed", combinedRunText, StringComparison.Ordinal);
            Assert.Contains(runs, run => run.Bold && string.Equals(run.Text, "no — nothing is failed", StringComparison.Ordinal));
        }

        [Fact]
        public void MarkdownToWord_PreferNarrativeSingleLineDefinitions_KeepsPlainNarrativeLineReadable() {
            const string markdown = "Interpretation: topology looks clean in this sample.";

            using var document = markdown.LoadFromMarkdown(new MarkdownToWordOptions {
                PreferNarrativeSingleLineDefinitions = true
            });
            var bodyText = string.Join("\n", document.Paragraphs.Select(p => p.Text));

            Assert.Contains("Interpretation: topology looks clean in this sample.", bodyText, StringComparison.Ordinal);
            Assert.DoesNotContain("Interpretation\\:", bodyText, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownToWord_PreferNarrativeSingleLineDefinitions_PreservesFormattedDefinitionRuns() {
            const string markdown = "Short answer: **no — nothing is failed** ✅";

            using var document = markdown.LoadFromMarkdown(new MarkdownToWordOptions {
                PreferNarrativeSingleLineDefinitions = true
            });
            var paragraph = document.Paragraphs.First(p => p.Text.Contains("Short answer", StringComparison.Ordinal));
            var runs = paragraph.GetRuns();
            var combinedRunText = string.Concat(runs.Select(run => run.Text));

            Assert.DoesNotContain("**", combinedRunText, StringComparison.Ordinal);
            Assert.Contains("no — nothing is failed", combinedRunText, StringComparison.Ordinal);
            Assert.Contains(runs, run => run.Bold && string.Equals(run.Text, "no — nothing is failed", StringComparison.Ordinal));
        }

        [Fact]
        public void MarkdownToWord_PreferNarrativeSingleLineDefinitions_Still_Renders_Grouped_Definition_List() {
            const string markdown = """
                Status: healthy
                Impact: none
                """;

            using var document = markdown.LoadFromMarkdown(new MarkdownToWordOptions {
                PreferNarrativeSingleLineDefinitions = true
            });
            var paragraphRunText = document.Paragraphs
                .Select(p => string.Concat(p.GetRuns().Select(run => run.Text)))
                .ToList();

            Assert.Contains("Status: healthy", paragraphRunText);
            Assert.Contains("Impact: none", paragraphRunText);
        }

        [Fact]
        public void MarkdownToWord_ReaderOptions_CanDisableOfficeIMOCallouts() {
            const string markdown = """
                > [!NOTE] Portable title
                > Still ordinary quoted text.
                """;

            using var officeDocument = markdown.LoadFromMarkdown(new MarkdownToWordOptions());
            var officeText = string.Join("\n", officeDocument.Paragraphs.Select(p => p.Text));

            Assert.Contains("Portable title", officeText, StringComparison.Ordinal);
            Assert.DoesNotContain("[!NOTE]", officeText, StringComparison.Ordinal);

            using var commonMarkDocument = markdown.LoadFromMarkdown(new MarkdownToWordOptions {
                ReaderOptions = OfficeIMO.Markdown.MarkdownReaderOptions.CreateCommonMarkProfile()
            });
            var commonMarkText = string.Join("\n", commonMarkDocument.Paragraphs.Select(p => p.Text));

            Assert.Contains("[!NOTE]", commonMarkText, StringComparison.Ordinal);
            Assert.Contains("Portable title", commonMarkText, StringComparison.Ordinal);
            Assert.Contains("Still ordinary quoted text.", commonMarkText, StringComparison.Ordinal);
        }
    }
}
