using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.Word.Markdown;
using OmdRenderer = OfficeIMO.MarkdownRenderer.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

[Collection("WordTests")]
public sealed class Markdown_CrossPipeline_Corpus_Tests {
    public sealed class CrossPipelineFixture {
        public CrossPipelineFixture(
            string name,
            string markdown,
            string[] markdownNeedles,
            string[] htmlNeedles,
            Func<MarkdownDoc>? createSourceDocument = null,
            Func<MarkdownReaderOptions>? createReaderOptions = null,
            Func<MarkdownRendererOptions>? createRendererOptions = null,
            Func<MarkdownToWordOptions>? createWordOptions = null,
            bool useSourceMarkdownAsPipelineInput = false) {
            Name = name;
            Markdown = markdown;
            MarkdownNeedles = markdownNeedles;
            HtmlNeedles = htmlNeedles;
            CreateSourceDocument = createSourceDocument;
            CreateReaderOptions = createReaderOptions;
            CreateRendererOptions = createRendererOptions;
            CreateWordOptions = createWordOptions;
            UseSourceMarkdownAsPipelineInput = useSourceMarkdownAsPipelineInput;
        }

        public string Name { get; }
        public string Markdown { get; }
        public string[] MarkdownNeedles { get; }
        public string[] HtmlNeedles { get; }
        public Func<MarkdownDoc>? CreateSourceDocument { get; }
        public Func<MarkdownReaderOptions>? CreateReaderOptions { get; }
        public Func<MarkdownRendererOptions>? CreateRendererOptions { get; }
        public Func<MarkdownToWordOptions>? CreateWordOptions { get; }
        public bool UseSourceMarkdownAsPipelineInput { get; }
    }

    public static IEnumerable<object[]> CrossPipelineFixtures() {
        yield return new object[] {
            new CrossPipelineFixture(
                name: "report-core-structure",
                markdown: """
# Report

Intro with **bold**, *italic*, `code`, and [Docs](https://example.com).

- **Replication:** healthy
- **FSMO:** technically OK

1. First
2. Second

| Name | Score |
| ---- | ----: |
| Alice | 98 |
| Bob | 91 |
""",
                markdownNeedles: new[] {
                    "# Report",
                    "**Replication:** healthy",
                    "**FSMO:** technically OK",
                    "[Docs](https://example.com)",
                    "| Name",
                    "| Score",
                    "1. First",
                    "2. Second"
                },
                htmlNeedles: new[] {
                    "<h1",
                    "<strong>bold</strong>",
                    "<em>italic</em>",
                    "<code>code</code>",
                    "<a href=\"https://example.com\">Docs</a>",
                    "<li><strong>Replication:</strong> healthy</li>",
                    "<li><strong>FSMO:</strong> technically OK</li>",
                    "<table>"
                })
        };

        yield return new object[] {
            new CrossPipelineFixture(
                name: "transcript-normalization-artifacts",
                markdown: """
-AD1
healthy for directory access

**Result
all 5 are healthy for directory access** with recommended LDAPS endpoints.
""",
                markdownNeedles: new[] {
                    "- AD1 healthy for directory access",
                    "**Result:** all 5 are healthy for directory access with recommended LDAPS endpoints."
                },
                htmlNeedles: new[] {
                    "<li>AD1 healthy for directory access</li>",
                    "<strong>Result:</strong>",
                    "all 5 are healthy for directory access with recommended LDAPS endpoints."
                },
                createReaderOptions: () => MarkdownTranscriptPreparation.CreateIntelligenceXTranscriptReaderOptions(
                    preservesGroupedDefinitionLikeParagraphs: false,
                    visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence),
                createRendererOptions: () => MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal(),
                createWordOptions: () => MarkdownToWordPresets.CreateIntelligenceXTranscript())
        };

        yield return new object[] {
            new CrossPipelineFixture(
                name: "inline-html-semantic-tags",
                markdown: """
Water H<sub>2</sub>O plus x<sup>2</sup> and <u>important</u>.
""",
                markdownNeedles: new[] {
                    "<sub>2</sub>",
                    "<sup>2</sup>",
                    "<u>important</u>"
                },
                htmlNeedles: new[] {
                    "<sub>2</sub>",
                    "<sup>2</sup>",
                    "<u>important</u>"
                },
                createSourceDocument: () => "<p>Water H<sub>2</sub>O plus x<sup>2</sup> and <u>important</u>.</p>".LoadFromHtml(),
                createRendererOptions: () => new MarkdownRendererOptions {
                    ReaderOptions = MarkdownReaderOptions.CreateOfficeIMOProfile(),
                    HtmlOptions = new HtmlOptions {
                        Style = HtmlStyle.Plain,
                        CssDelivery = CssDelivery.None,
                        BodyClass = null,
                        Kind = HtmlKind.Fragment
                    }
                },
                useSourceMarkdownAsPipelineInput: true)
        };
    }

    [Theory]
    [MemberData(nameof(CrossPipelineFixtures))]
    public void Markdown_CrossPipeline_Corpus_Retains_Core_Structure(
        CrossPipelineFixture fixture) {
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            Kind = HtmlKind.Fragment
        };

        var readerOptions = fixture.CreateReaderOptions?.Invoke() ?? MarkdownReaderOptions.CreateOfficeIMOProfile();
        var source = fixture.CreateSourceDocument?.Invoke() ?? MarkdownReader.Parse(fixture.Markdown, readerOptions);
        var sourceMarkdown = NormalizeMarkdown(source.ToMarkdown());
        var sourceHtml = source.ToHtmlFragment(htmlOptions);
        var pipelineMarkdown = fixture.UseSourceMarkdownAsPipelineInput ? sourceMarkdown : fixture.Markdown;

        var rendererOptions = fixture.CreateRendererOptions?.Invoke() ?? new MarkdownRendererOptions {
            HtmlOptions = new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                Kind = HtmlKind.Fragment
            }
        };
        var renderedHtml = OmdRenderer.RenderBodyHtml(pipelineMarkdown, rendererOptions);

        var htmlRoundTrip = sourceHtml.LoadFromHtml(new HtmlToMarkdownOptions());
        var htmlRoundTripMarkdown = NormalizeMarkdown(htmlRoundTrip.ToMarkdown());

        using var wordDocument = pipelineMarkdown.LoadFromMarkdown(
            fixture.CreateWordOptions?.Invoke() ?? new MarkdownToWordOptions { FontFamily = "Calibri" });
        var wordRoundTripMarkdown = NormalizeMarkdown(
            wordDocument.ToMarkdown(new WordToMarkdownOptions { EnableUnderline = true }));

        foreach (var markdownNeedle in fixture.MarkdownNeedles) {
            Assert.Contains(markdownNeedle, sourceMarkdown, StringComparison.Ordinal);
            Assert.Contains(markdownNeedle, htmlRoundTripMarkdown, StringComparison.Ordinal);
            Assert.Contains(markdownNeedle, wordRoundTripMarkdown, StringComparison.Ordinal);
        }

        foreach (var htmlNeedle in fixture.HtmlNeedles) {
            Assert.Contains(htmlNeedle, sourceHtml, StringComparison.Ordinal);
            Assert.Contains(htmlNeedle, renderedHtml, StringComparison.Ordinal);
        }
    }

    private static string NormalizeMarkdown(string markdown) {
        return (markdown ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim();
    }
}
