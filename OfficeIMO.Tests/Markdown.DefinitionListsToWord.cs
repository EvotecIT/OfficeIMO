using System;
using System.Linq;
using OfficeIMO.Markdown;
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

        [Fact]
        public void MarkdownToWordPreset_IntelligenceXTranscript_ConfiguresTypedTranscriptDefaults() {
            var options = MarkdownToWordPresets.CreateIntelligenceXTranscript(
                ["C:\\allowed-a", "", "C:\\allowed-a", "C:\\allowed-b"],
                2500);

            Assert.Equal("Calibri", options.FontFamily);
            Assert.True(options.AllowLocalImages);
            Assert.True(options.PreferNarrativeSingleLineDefinitions);
            Assert.True(options.FitImagesToPageContentWidth);
            Assert.False(options.FitImagesToContextWidth);
            Assert.Equal(100d, options.MaxImageWidthPercentOfContent);
            Assert.Equal(2000, options.MaxImageWidthPixels);
            Assert.Equal(2, options.AllowedImageDirectories.Count);
            Assert.Contains("C:\\allowed-a", options.AllowedImageDirectories);
            Assert.Contains("C:\\allowed-b", options.AllowedImageDirectories);
            Assert.NotNull(options.ReaderOptions);
            Assert.IsType<MarkdownReaderOptions>(options.ReaderOptions);
            Assert.True(options.ReaderOptions!.PreferNarrativeSingleLineDefinitions);
            Assert.True(options.ReaderOptions.Callouts);
            Assert.True(options.ReaderOptions.DefinitionLists);
            Assert.NotNull(options.ReaderOptions.InputNormalization);
            Assert.Contains(options.ReaderOptions.DocumentTransforms, transform => transform is MarkdownSimpleDefinitionListParagraphTransform);
            Assert.Contains(options.ReaderOptions.DocumentTransforms, transform => transform is MarkdownJsonVisualCodeBlockTransform);
        }

        [Fact]
        public void MarkdownToWordPreset_IntelligenceXTranscript_Flattens_Simple_Grouped_Definitions_In_ReaderPipeline() {
            const string markdown = """
                Status: healthy
                Impact: none
                """;

            var options = MarkdownToWordPresets.CreateIntelligenceXTranscript();
            var document = MarkdownReader.Parse(markdown, options.ReaderOptions);

            Assert.Collection(document.Blocks,
                block => {
                    var paragraph = Assert.IsType<ParagraphBlock>(block);
                    Assert.Equal("Status: healthy", paragraph.Inlines.RenderMarkdown());
                },
                block => {
                    var paragraph = Assert.IsType<ParagraphBlock>(block);
                    Assert.Equal("Impact: none", paragraph.Inlines.RenderMarkdown());
                });
        }

        [Fact]
        public void MarkdownToWordCapabilities_DetectNarrativeSingleLineDefinitionSupport() {
            const string markdown = """
                Status: healthy
                Impact: none
                """;

            using var document = markdown.LoadFromMarkdown(
                MarkdownToWordPresets.CreateIntelligenceXTranscript());
            var bodyParagraphs = document.Paragraphs
                .Select(p => string.Concat(p.GetRuns().Select(run => run.Text)))
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            var actualSupport = bodyParagraphs.Contains("Status: healthy", StringComparer.Ordinal)
                                && bodyParagraphs.Contains("Impact: none", StringComparer.Ordinal);

            Assert.Equal(
                actualSupport,
                MarkdownToWordCapabilities.PreservesNarrativeSingleLineDefinitionsAsSeparateParagraphs());
        }
    }
}
