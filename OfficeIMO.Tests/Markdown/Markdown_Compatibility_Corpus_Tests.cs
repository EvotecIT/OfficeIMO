using System;
using System.IO;
using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Compatibility_Corpus_Tests {
    [Fact]
    public void Compatibility_Fixture_Profiles_Keep_OfficeIMO_Extensions_Explicit() {
        string markdown = LoadCompatibilityFixture("portable-profile-boundary.md");
        var htmlOptions = CreatePlainHtmlOptions();

        string officeHtml = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateOfficeIMOProfile()).ToHtmlFragment(htmlOptions);
        string portableHtml = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreatePortableProfile()).ToHtmlFragment(htmlOptions);

        Assert.Contains("class=\"callout", officeHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("contains-task-list", officeHtml, StringComparison.Ordinal);
        Assert.Contains("task-list-item-checkbox", officeHtml, StringComparison.Ordinal);
        Assert.Contains("class=\"footnotes\"", officeHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("href=\"https://example.com/path_(x)\"", officeHtml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("[TOC]", officeHtml, StringComparison.Ordinal);

        Assert.DoesNotContain("class=\"callout", portableHtml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("contains-task-list", portableHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("task-list-item-checkbox", portableHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"footnotes\"", portableHtml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("href=\"https://example.com/path_(x)\"", portableHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("[TOC]", portableHtml, StringComparison.Ordinal);
        Assert.Contains("[!NOTE]", portableHtml, StringComparison.Ordinal);
        Assert.Contains("[ ] Track parser parity", portableHtml, StringComparison.Ordinal);
        Assert.Contains("[x] Keep AST honest", portableHtml, StringComparison.Ordinal);
        Assert.Contains("[^shape]", portableHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void Compatibility_Fixture_Ix_Aliases_Remain_OptIn_For_Renderer_Hosts() {
        string markdown = LoadCompatibilityFixture("ix-visuals.md");

        var generic = MarkdownRendererPresets.CreateStrictMinimal();
        generic.Chart.Enabled = true;
        generic.Network.Enabled = true;

        var ix = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        ix.Chart.Enabled = true;
        ix.Network.Enabled = true;

        string genericHtml = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, generic);
        string ixHtml = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, ix);

        Assert.Contains("language-ix-chart", genericHtml, StringComparison.Ordinal);
        Assert.Contains("language-ix-network", genericHtml, StringComparison.Ordinal);
        Assert.Contains("language-ix-dataview", genericHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"omd-visual omd-chart\"", genericHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"omd-visual omd-network\"", genericHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"omd-visual omd-dataview\"", genericHtml, StringComparison.Ordinal);

        Assert.Contains("class=\"omd-visual omd-chart\"", ixHtml, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-visual omd-network\"", ixHtml, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-visual omd-dataview\"", ixHtml, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"ix-chart\"", ixHtml, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"ix-network\"", ixHtml, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"ix-dataview\"", ixHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void Compatibility_Fixture_Html_Ingestion_Preserves_Rich_Ast_Before_Serialization() {
        string html = LoadCompatibilityFixture("html-rich-ast.html");

        MarkdownDoc document = html.LoadFromHtml();
        string portableMarkdown = html.ToMarkdown(HtmlToMarkdownOptions.CreatePortableProfile());
        string renderedHtml = document.ToHtmlFragment(CreatePlainHtmlOptions());

        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Level == 1 && heading.Text == "HTML Corpus");

        var table = Assert.Single(document.Blocks.OfType<TableBlock>());
        Assert.Collection(table.RowCells[0][1].Blocks,
            block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.IsType<QuoteBlock>(block));

        var definitions = Assert.Single(document.Blocks.OfType<DefinitionListBlock>());
        Assert.Equal(2, definitions.Entries.Count);
        Assert.Equal("Term A", definitions.Entries[0].Term.RenderMarkdown());
        Assert.Equal("Term B", definitions.Entries[1].Term.RenderMarkdown());
        Assert.Collection(definitions.Entries[0].DefinitionBlocks,
            block => Assert.Equal("First paragraph", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.Equal("Second paragraph", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

        Assert.Contains("<td><p>Intro</p><blockquote><p>Quoted</p></blockquote></td>", renderedHtml, StringComparison.Ordinal);
        Assert.Contains("Intro", portableMarkdown, StringComparison.Ordinal);
        Assert.Contains("Quoted", portableMarkdown, StringComparison.Ordinal);
        Assert.Contains("First paragraph", portableMarkdown, StringComparison.Ordinal);
        Assert.Contains("Second paragraph", portableMarkdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Compatibility_Fixture_SharedVisualHosts_Html_Ingestion_Preserves_Generic_Semantic_Block_Recovery() {
        string html = LoadHtmlFixture("shared-visual-hosts.html");

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/visuals/archive.html")
        });
        string markdown = document.ToMarkdown(MarkdownWriteOptions.CreateOfficeIMOProfile());

        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Level == 1 && heading.Text == "Shared Visual Archive");

        var semanticBlocks = document.Blocks.OfType<SemanticFencedBlock>().ToList();
        Assert.Equal(3, semanticBlocks.Count);
        Assert.Contains(semanticBlocks, block => block.Language == "chart" && block.Caption == "Chart preview");
        Assert.Contains(semanticBlocks, block => block.Language == "network" && block.Caption == "Network preview");
        Assert.Contains(semanticBlocks, block => block.Language == "dataview" && block.Caption == "Dataview preview");
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"label\":\"Count\"", StringComparison.Ordinal));
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"label\":\"memberOf\"", StringComparison.Ordinal));
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"kind\":\"ix_tool_dataview_v1\"", StringComparison.Ordinal));

        Assert.Contains("```chart", markdown, StringComparison.Ordinal);
        Assert.Contains("```network", markdown, StringComparison.Ordinal);
        Assert.Contains("```dataview", markdown, StringComparison.Ordinal);
        Assert.Contains("_Chart preview_", markdown, StringComparison.Ordinal);
        Assert.Contains("_Network preview_", markdown, StringComparison.Ordinal);
        Assert.Contains("_Dataview preview_", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("data-omd-visual-kind", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Compatibility_Fixture_Ix_CompatibilityTranscriptVisualPack_Preserves_Semantic_Block_Recovery() {
        string markdown = LoadCompatibilityFixture("ix-compat-transcript-visual-pack.md");

        var ix = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        ix.Chart.Enabled = true;
        ix.Network.Enabled = true;
        ix.Mermaid.Enabled = true;

        string html = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, ix);
        MarkdownDoc document = html.LoadFromHtml();

        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "Assistant (20:30: 13)");
        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "Assistant (20:36: 24)");
        Assert.True(document.Blocks.OfType<TableBlock>().Count() >= 1);
        Assert.True(document.Blocks.OfType<OrderedListBlock>().Count() >= 1);

        var semanticBlocks = document.Blocks.OfType<SemanticFencedBlock>().ToList();
        Assert.Equal(5, semanticBlocks.Count(block => block.Language == "ix-chart"));
        Assert.Equal(2, semanticBlocks.Count(block => block.Language == "mermaid"));
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"label\": \"Broken\"", StringComparison.Ordinal));
        Assert.Contains("## Proactive checks", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void Compatibility_Fixture_Ix_CompatibilityTranscriptChartSuite_Preserves_ChartHeavy_Transcript_Recovery() {
        string markdown = LoadCompatibilityFixture("ix-compat-transcript-chart-suite.md");

        var ix = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        ix.Chart.Enabled = true;
        ix.Mermaid.Enabled = true;

        string html = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, ix);
        MarkdownDoc document = html.LoadFromHtml();

        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "Assistant (20:36: 24)");
        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "Assistant (20:37: 49)");
        Assert.True(document.Blocks.OfType<TableBlock>().Count() >= 1);

        var semanticBlocks = document.Blocks.OfType<SemanticFencedBlock>().ToList();
        Assert.Equal(4, semanticBlocks.Count(block => block.Language == "ix-chart"));
        Assert.Equal(1, semanticBlocks.Count(block => block.Language == "mermaid"));
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"label\": \"Critical\"", StringComparison.Ordinal));
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"label\": \"Privileged Changes\"", StringComparison.Ordinal));
        Assert.Contains("Risk distribution", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void Compatibility_Fixture_Ix_PortableExportLegacyJsonVisuals_Separates_Compatibility_And_Portable_Export_Lanes() {
        string markdown = LoadCompatibilityFixture("ix-portable-export-legacy-json-visuals.md");

        string ixCompatibilityExport = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForExport(markdown);
        string portableExport = MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForExport(
            markdown,
            MarkdownVisualFenceLanguageMode.GenericSemanticFence);

        Assert.Contains("```ix-chart", ixCompatibilityExport, StringComparison.Ordinal);
        Assert.Contains("```ix-dataview", ixCompatibilityExport, StringComparison.Ordinal);
        Assert.DoesNotContain("```chart", ixCompatibilityExport, StringComparison.Ordinal);
        Assert.DoesNotContain("```dataview", ixCompatibilityExport, StringComparison.Ordinal);

        Assert.Contains("```chart", portableExport, StringComparison.Ordinal);
        Assert.Contains("```dataview", portableExport, StringComparison.Ordinal);
        Assert.DoesNotContain("```ix-chart", portableExport, StringComparison.Ordinal);
        Assert.DoesNotContain("```ix-dataview", portableExport, StringComparison.Ordinal);
        Assert.DoesNotContain("ix:cached-tool-evidence:v1", portableExport, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\"label\": \"Count\"", portableExport, StringComparison.Ordinal);
    }

    private static HtmlOptions CreatePlainHtmlOptions() {
        return new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        };
    }

    private static string LoadCompatibilityFixture(string name) {
        string path = Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..",
            "Markdown",
            "Fixtures",
            "Compatibility",
            name);

        return File.ReadAllText(path);
    }

    private static string LoadHtmlFixture(string name) {
        string path = Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..",
            "Markdown",
            "Fixtures",
            name);

        return File.ReadAllText(path);
    }
}

