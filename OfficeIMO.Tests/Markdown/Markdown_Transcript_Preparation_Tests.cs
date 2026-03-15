using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownTranscriptPreparationTests {
    [Fact]
    public void PrepareIntelligenceXTranscriptBody_ComposesNormalizationAndOrderedListRepair() {
        const string markdown = """
            1) First check
            2) Second check
            """;

        var prepared = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptBody(markdown)
            .Replace("\r\n", "\n");

        Assert.Contains("1. First check\n\n2. Second check", prepared, StringComparison.Ordinal);
    }

    [Fact]
    public void CreateIntelligenceXTranscriptReaderOptions_Composes_Transcript_Normalization_And_Optional_Document_Transforms() {
        var preserved = MarkdownTranscriptPreparation.CreateIntelligenceXTranscriptReaderOptions(
            preservesGroupedDefinitionLikeParagraphs: true);
        var flattened = MarkdownTranscriptPreparation.CreateIntelligenceXTranscriptReaderOptions(
            preservesGroupedDefinitionLikeParagraphs: false,
            visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);

        Assert.NotNull(preserved.InputNormalization);
        Assert.True(preserved.InputNormalization!.NormalizeCollapsedOrderedListBoundaries);
        Assert.True(preserved.PreferNarrativeSingleLineDefinitions);
        Assert.DoesNotContain(preserved.DocumentTransforms, transform => transform is MarkdownSimpleDefinitionListParagraphTransform);
        Assert.DoesNotContain(preserved.DocumentTransforms, transform => transform is MarkdownJsonVisualCodeBlockTransform);

        Assert.NotNull(flattened.InputNormalization);
        Assert.True(flattened.InputNormalization!.NormalizeCollapsedOrderedListBoundaries);
        Assert.True(flattened.PreferNarrativeSingleLineDefinitions);
        Assert.Contains(flattened.DocumentTransforms, transform => transform is MarkdownSimpleDefinitionListParagraphTransform);
        Assert.Contains(flattened.DocumentTransforms, transform =>
            transform is MarkdownJsonVisualCodeBlockTransform visual
            && visual.LanguageMode == MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
    }

    [Fact]
    public void PrepareIntelligenceXTranscriptForExport_CollapsesDuplicateBlankLines() {
        const string markdown = """
            # Transcript


            Status: healthy



            ### Result
            """;

        var prepared = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForExport(markdown)
            .Replace("\r\n", "\n");

        Assert.DoesNotContain("\n\n\n", prepared, StringComparison.Ordinal);
        Assert.Contains("# Transcript\n\nStatus: healthy\n\n### Result", prepared, StringComparison.Ordinal);
    }

    [Fact]
    public void PrepareIntelligenceXTranscriptDocument_Can_Parse_Transcript_Artifacts_Via_Shared_Reader_Contract() {
        const string markdown = """
            1) First check
            2) Second check
            """;

        var document = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptDocument(markdown);
        var list = Assert.IsType<OrderedListBlock>(Assert.Single(document.Blocks));

        Assert.Equal(2, list.Items.Count);
        Assert.Equal("First check", list.Items[0].Content.RenderMarkdown());
        Assert.Equal("Second check", list.Items[1].Content.RenderMarkdown());
    }

    [Fact]
    public void PrepareIntelligenceXTranscriptDocument_Can_Upgrade_Legacy_Visual_Json_Via_Shared_Reader_Contract() {
        const string markdown = """
            ```json
            {"type":"bar","data":{"labels":["A"],"datasets":[{"label":"Count","data":[1]}]}}
            ```
            """;

        var document = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptDocument(
            markdown,
            visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);
        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));

        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("ix-chart", block.Language);
    }

    [Fact]
    public void PrepareIntelligenceXTranscriptForDocx_OptionallySeparatesGroupedDefinitionLikeParagraphs() {
        const string markdown = """
            Status: healthy
            Impact: none
            """;

        var preserved = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForDocx(markdown, preservesGroupedDefinitionLikeParagraphs: true)
            .Replace("\r\n", "\n");
        var repaired = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForDocx(markdown, preservesGroupedDefinitionLikeParagraphs: false)
            .Replace("\r\n", "\n");

        Assert.Contains("Status: healthy\nImpact: none", preserved, StringComparison.Ordinal);
        Assert.Contains("Status: healthy\n\nImpact: none", repaired, StringComparison.Ordinal);
    }

    [Fact]
    public void PrepareIntelligenceXTranscriptDocumentForDocx_Optionally_Flattens_Grouped_Definitions_Via_Ast() {
        const string markdown = """
            Status: healthy
            Impact: none
            """;

        var preserved = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptDocumentForDocx(
            markdown,
            preservesGroupedDefinitionLikeParagraphs: true);
        var repaired = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptDocumentForDocx(
            markdown,
            preservesGroupedDefinitionLikeParagraphs: false);

        Assert.IsType<DefinitionListBlock>(Assert.Single(preserved.Blocks));
        Assert.Collection(repaired.Blocks,
            block => Assert.Equal("Status: healthy", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.Equal("Impact: none", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));
    }

    [Fact]
    public void PrepareIntelligenceXTranscriptDocumentForDocx_Can_Compose_Definition_Compatibility_And_Visual_Upgrade() {
        const string markdown = """
            Status: healthy
            Impact: none

            ```json
            {"type":"bar","data":{"labels":["A"],"datasets":[{"label":"Count","data":[1]}]}}
            ```
            """;

        var document = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptDocumentForDocx(
            markdown,
            preservesGroupedDefinitionLikeParagraphs: false,
            visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);

        Assert.Collection(document.Blocks,
            block => Assert.Equal("Status: healthy", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.Equal("Impact: none", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => {
                var visual = Assert.IsType<SemanticFencedBlock>(block);
                Assert.Equal(MarkdownSemanticKinds.Chart, visual.SemanticKind);
                Assert.Equal("ix-chart", visual.Language);
            });
    }
}
