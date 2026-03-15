using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using Xunit;
using MarkdownRendererShell = OfficeIMO.MarkdownRenderer.MarkdownRenderer;
using System.IO;

namespace OfficeIMO.Tests;

public sealed class MarkdownHtmlToMarkdownTests {
    [Fact]
    public void HtmlToMarkdown_Converter_UsesDefaultOptionsWhenNull() {
        var converter = new HtmlToMarkdownConverter();

        string markdown = converter.Convert("<p>Hello</p>", options: null);
        MarkdownDoc document = converter.ConvertToDocument("<p>Hello</p>", options: null);

        Assert.Contains("Hello", markdown, StringComparison.Ordinal);
        Assert.Single(document.Blocks);
        Assert.IsType<ParagraphBlock>(document.Blocks[0]);
    }

    [Fact]
    public void HtmlToMarkdown_ConvertsCommonDocumentBlocks() {
        string html = "<html><body><h1>Hello</h1><p>A <strong>bold</strong> <a href=\"https://example.com\">link</a>.</p><ul><li>One</li><li>Two</li></ul></body></html>";

        string markdown = html.ToMarkdown();

        Assert.Contains("# Hello", markdown, StringComparison.Ordinal);
        Assert.Contains("**bold**", markdown, StringComparison.Ordinal);
        Assert.Contains("[link](https://example.com)", markdown, StringComparison.Ordinal);
        Assert.Contains("- One", markdown, StringComparison.Ordinal);
        Assert.Contains("- Two", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Trims_StrongBoundaryWhitespace_Before_InlineParsing() {
        string html = "<p><strong> LDAP/Kerberos health on all DCs </strong> next</p>";

        MarkdownDoc document = html.LoadFromHtml();
        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("<strong>LDAP/Kerberos health on all DCs</strong> next", renderedHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("** LDAP/Kerberos health on all DCs **", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Convert_Uses_Configured_MarkdownWriteOptions() {
        var options = new HtmlToMarkdownOptions {
            MarkdownWriteOptions = new MarkdownWriteOptions()
        };
        options.MarkdownWriteOptions.BlockRenderExtensions.Add(new MarkdownBlockMarkdownRenderExtension(
            "Test.Paragraph.Override",
            typeof(ParagraphBlock),
            (block, _) => block is ParagraphBlock ? "PARAGRAPH-OVERRIDE" : null));

        string markdown = "<p>Hello</p>".ToMarkdown(options);

        Assert.Equal("PARAGRAPH-OVERRIDE", markdown.Trim());
    }

    [Fact]
    public void HtmlToMarkdown_LoadFromHtml_ProducesTypedBlocks() {
        string html = "<html><body><h2>Section</h2><blockquote><p>Quoted</p></blockquote><details open><summary>More</summary><p>Hidden text</p></details></body></html>";

        MarkdownDoc document = html.LoadFromHtml();

        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Level == 2 && heading.Text == "Section");
        Assert.Contains(document.Blocks, block => block is QuoteBlock);
        Assert.Contains(document.Blocks, block => block is DetailsBlock details && details.Open);
    }

    [Fact]
    public void HtmlToMarkdown_Enforces_MaxInputCharacters() {
        var options = new HtmlToMarkdownOptions {
            MaxInputCharacters = 12
        };

        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => "<p>0123456789</p>".LoadFromHtml(options));
        Assert.Contains("MaxInputCharacters", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_ConvertsHtmlFragmentWithoutBodyWrapper() {
        string html = "<h2>Fragment</h2><p>Body</p>";

        MarkdownDoc document = html.LoadFromHtml();

        Assert.Collection(document.Blocks,
            block => Assert.IsType<HeadingBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_ResolvesRelativeLinksWithBaseUri() {
        string html = "<p><a href=\"guide/start\">Docs</a></p>";

        string markdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/docs/")
        });

        Assert.Contains("[Docs](https://example.com/docs/guide/start)", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedBlocks_WhenRequested() {
        string html = "<custom-widget data-name=\"demo\">Hello</custom-widget>";

        string markdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = true
        });

        Assert.Contains("<custom-widget", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedInlineHtml_WhenRequested() {
        string html = "<p>Hello <custom-inline data-name=\"demo\">world</custom-inline></p>";

        string markdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            PreserveUnsupportedInlineHtml = true
        });

        Assert.Contains("<custom-inline", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedBlockElementsBetweenParagraphs() {
        string html = "<p>Alpha</p><custom-widget data-name=\"demo\">Hello</custom-widget><p>Omega</p>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = true
        });

        Assert.Collection(document.Blocks,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<HtmlRawBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));

        string markdown = document.ToMarkdown();
        int alphaIndex = markdown.IndexOf("Alpha", StringComparison.Ordinal);
        int widgetIndex = markdown.IndexOf("<custom-widget", StringComparison.Ordinal);
        int omegaIndex = markdown.IndexOf("Omega", StringComparison.Ordinal);
        Assert.True(alphaIndex >= 0, "Expected Alpha in markdown output.");
        Assert.True(widgetIndex > alphaIndex, "Expected raw custom block after the opening paragraph.");
        Assert.True(omegaIndex > widgetIndex, "Expected trailing paragraph after the custom block.");
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedBlockElementsInsideListItems() {
        string html = "<ul><li><p>Alpha</p><custom-widget data-name=\"demo\">Hello</custom-widget><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = true
        });

        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<HtmlRawBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedBlockElementsInsideSections() {
        string html = "<section><p>Alpha</p><custom-widget data-name=\"demo\">Hello</custom-widget><p>Omega</p></section>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = true
        });

        Assert.Collection(document.Blocks,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<HtmlRawBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesUnsupportedBlockElementsInsideDetails() {
        string html = "<details open><summary>More</summary><p>Alpha</p><custom-widget data-name=\"demo\">Hello</custom-widget><p>Omega</p></details>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = true
        });

        var details = Assert.IsType<DetailsBlock>(Assert.Single(document.Blocks));
        Assert.NotNull(details.Summary);
        Assert.Collection(details.ChildBlocks,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<HtmlRawBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_CapturesFigureCaptionOnImageBlocks() {
        string html = "<figure><img src=\"/img/demo.png\" alt=\"Demo\" /><figcaption>Example caption</figcaption></figure>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        Assert.Equal("https://example.com/img/demo.png", image.Path);
        Assert.Equal("Demo", image.Alt);
        Assert.Equal("Example caption", image.Caption);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesFigureOrderInsideListItems() {
        string html = "<ul><li><p>Alpha</p><figure><img src=\"/img/demo.png\" alt=\"Demo\" /><figcaption>Caption</figcaption></figure><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/")
        });

        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => {
                var image = Assert.IsType<ImageBlock>(block);
                Assert.Equal("https://example.com/img/demo.png", image.Path);
                Assert.Equal("Caption", image.Caption);
            },
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesParagraphBreaksInsideTableCells() {
        string html = "<table><tr><th>Section</th><th>Notes</th></tr><tr><td>Alpha</td><td><p>First</p><p>Second</p></td></tr></table>";

        MarkdownDoc document = html.LoadFromHtml();
        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));

        Assert.Equal("Alpha", table.Rows[0][0]);
        Assert.Equal("First\n\nSecond", table.Rows[0][1]);
        Assert.Collection(table.RowCells[0][1].Blocks,
            block => Assert.Equal("First", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.Equal("Second", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

        string markdown = document.ToMarkdown();
        Assert.Contains("First<br><br>Second", markdown, StringComparison.Ordinal);

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<td><p>First</p><p>Second</p></td>", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMixedBlockContentInsideTableCells() {
        string html = "<table><tr><th>Section</th><th>Notes</th></tr><tr><td>Alpha</td><td><p>Intro</p><blockquote><p>Quoted</p></blockquote></td></tr></table>";

        MarkdownDoc document = html.LoadFromHtml();
        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));

        Assert.Collection(table.RowCells[0][1].Blocks,
            block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.IsType<QuoteBlock>(block));

        string renderedHtml = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<td><p>Intro</p><blockquote><p>Quoted</p></blockquote></td>", renderedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMixedListItemBlockOrder() {
        string html = "<ul><li><p>Alpha</p><blockquote><p>Quoted</p></blockquote><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<QuoteBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));

        string markdown = document.ToMarkdown();
        int alphaIndex = markdown.IndexOf("Alpha", StringComparison.Ordinal);
        int quoteIndex = markdown.IndexOf("Quoted", StringComparison.Ordinal);
        int omegaIndex = markdown.IndexOf("Omega", StringComparison.Ordinal);
        Assert.True(alphaIndex >= 0, "Expected Alpha in markdown output.");
        Assert.True(quoteIndex > alphaIndex, "Expected quoted content after the opening paragraph.");
        Assert.True(omegaIndex > quoteIndex, "Expected trailing paragraph after the quote block.");
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMultipleDefinitionsPerTerm() {
        string html = "<dl><dt>Term</dt><dd>First definition</dd><dd>Second definition</dd></dl>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));

        Assert.Equal(2, list.Entries.Count);
        Assert.Equal("Term", list.Entries[0].Term.RenderMarkdown());
        Assert.Equal("First definition", Assert.IsType<ParagraphBlock>(Assert.Single(list.Entries[0].DefinitionBlocks)).Inlines.RenderMarkdown());
        Assert.Equal("Term", list.Entries[1].Term.RenderMarkdown());
        Assert.Equal("Second definition", Assert.IsType<ParagraphBlock>(Assert.Single(list.Entries[1].DefinitionBlocks)).Inlines.RenderMarkdown());
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMultipleParagraphsInDefinitionValues() {
        string html = "<dl><dt>Term</dt><dd><p>First paragraph</p><p>Second paragraph</p></dd></dl>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));

        Assert.Single(list.Entries);
        Assert.Equal("Term", list.Entries[0].Term.RenderMarkdown());
        Assert.Collection(list.Entries[0].DefinitionBlocks,
            block => Assert.Equal("First paragraph", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.Equal("Second paragraph", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

        string markdown = document.ToMarkdown();
        Assert.Contains("Term: First paragraph", markdown, StringComparison.Ordinal);
        Assert.Contains("Second paragraph", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMixedBlockContentInDefinitionValues() {
        string html = "<dl><dt>Term</dt><dd><p>Intro</p><blockquote><p>Quoted</p></blockquote><ul><li>Nested</li></ul></dd></dl>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));

        Assert.Single(list.Entries);
        Assert.Equal("Term", list.Entries[0].Term.RenderMarkdown());
        Assert.Collection(list.Entries[0].DefinitionBlocks,
            block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
            block => Assert.IsType<QuoteBlock>(block),
            block => Assert.IsType<UnorderedListBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesNestedListOrderInsideListItem() {
        string html = "<ul><li><p>Alpha</p><ul><li>Nested</li></ul><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<UnorderedListBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));

        string markdown = document.ToMarkdown();
        int alphaIndex = markdown.IndexOf("Alpha", StringComparison.Ordinal);
        int nestedIndex = markdown.IndexOf("Nested", StringComparison.Ordinal);
        int omegaIndex = markdown.IndexOf("Omega", StringComparison.Ordinal);
        Assert.True(alphaIndex >= 0, "Expected Alpha in markdown output.");
        Assert.True(nestedIndex > alphaIndex, "Expected nested list content after the opening paragraph.");
        Assert.True(omegaIndex > nestedIndex, "Expected trailing paragraph after the nested list.");
    }

    [Fact]
    public void HtmlToMarkdown_PreservesDetailsOrderInsideListItem() {
        string html = "<ul><li><p>Alpha</p><details open><summary>More</summary><p>Hidden</p></details><p>Omega</p></li></ul>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Collection(item.BlockChildren,
            block => Assert.IsType<ParagraphBlock>(block),
            block => Assert.IsType<DetailsBlock>(block),
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void HtmlToMarkdown_PreservesMultipleTermsPerDefinitionGroup() {
        string html = "<dl><dt>Alpha</dt><dt>Beta</dt><dd>Shared definition</dd><dd>Follow-up definition</dd></dl>";

        MarkdownDoc document = html.LoadFromHtml();
        var list = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));

        Assert.Equal(4, list.Entries.Count);
        Assert.Equal(("Alpha", "Shared definition"), list.Items[0]);
        Assert.Equal(("Beta", "Shared definition"), list.Items[1]);
        Assert.Equal(("Alpha", "Follow-up definition"), list.Items[2]);
        Assert.Equal(("Beta", "Follow-up definition"), list.Items[3]);
        Assert.Equal("Alpha", list.Entries[0].Term.RenderMarkdown());
        Assert.Equal("Shared definition", Assert.IsType<ParagraphBlock>(Assert.Single(list.Entries[0].DefinitionBlocks)).Inlines.RenderMarkdown());
    }

    [Fact]
    public void HtmlToMarkdown_ConvertsSharedVisualHostIntoSemanticFencedBlock() {
        var payload = MarkdownVisualContract.CreatePayload("{\"type\":\"bar\"}");
        string html = MarkdownVisualContract.BuildElementHtml(
            "div",
            "omd-visual omd-custom",
            MarkdownSemanticKinds.Chart,
            "vendor-chart",
            payload);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("vendor-chart", block.Language);
        Assert.Equal("{\"type\":\"bar\"}", block.Content);
        Assert.Equal(
            NormalizeMarkdown("```vendor-chart\n{\"type\":\"bar\"}\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_RenderedIxDataviewHtml_BackToSemanticFence() {
        const string raw = """
{"title":"Replication Summary","summary":"Latest replication posture","kind":"ix_tool_dataview_v1","call_id":"call_123","headers":["Domain","Status"],"items":[{"Domain":"ad.evotec.xyz","Status":"Healthy"}]}
""";
        var options = MarkdownRendererPresets.CreateStrictMinimal();
        MarkdownRendererIntelligenceXAdapter.Apply(options);
        MarkdownRendererIntelligenceXLegacyMigration.Apply(options);
        string html = MarkdownRendererShell.RenderBodyHtml("```ix-dataview\n" + raw + "\n```", options);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.DataView, block.SemanticKind);
        Assert.Equal("ix-dataview", block.Language);
        Assert.Equal(raw, block.Content);
        Assert.Equal(
            NormalizeMarkdown("```ix-dataview\n" + raw + "\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_RenderedIxNetworkHtml_BackToSemanticFence() {
        const string raw = """
{"nodes":[{"id":"A","label":"User"},{"id":"B","label":"Group"}],"edges":[{"from":"A","to":"B","label":"memberOf"}]}
""";
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Network.Enabled = true;
        string html = MarkdownRendererShell.RenderBodyHtml("```ix-network\n" + raw + "\n```", options);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Network, block.SemanticKind);
        Assert.Equal("ix-network", block.Language);
        Assert.Equal(raw, block.Content);
        Assert.Equal(
            NormalizeMarkdown("```ix-network\n" + raw + "\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_RenderedMermaidHtml_BackToSemanticFence() {
        const string raw = """
flowchart LR
A[Markdown Input] --> B{Parser OK?}
B --> C[Render Mermaid]
""";
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Mermaid.Enabled = true;
        string html = MarkdownRendererShell.RenderBodyHtml("```mermaid\n" + raw + "\n```", options);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Mermaid, block.SemanticKind);
        Assert.Equal("mermaid", block.Language);
        Assert.Equal(NormalizeMarkdown(raw), NormalizeMarkdown(block.Content));
        Assert.Equal(
            NormalizeMarkdown("```mermaid\n" + raw + "\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_RenderedMathHtml_BackToSemanticFence() {
        const string raw = """
\sum_{i=1}^{n} x_i^2
""";
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Math.Enabled = true;
        string html = MarkdownRendererShell.RenderBodyHtml("```math\n" + raw + "\n```", options);

        MarkdownDoc document = html.LoadFromHtml();

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Math, block.SemanticKind);
        Assert.Equal("math", block.Language);
        Assert.Equal(raw, block.Content);
        Assert.Equal(
            NormalizeMarkdown("```math\n" + raw + "\n```"),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_MixedTranscriptFragment_WithVisuals_AndStructuredBlocks() {
        const string markdown = """
# Transcript Sample

## Proactive checks

| Check | Expected |
|---|---|
| Fence integrity | Visuals render |

```ix-chart
{"type":"bar","data":{"labels":["A"],"datasets":[{"label":"Count","data":[1]}]}}
```

```ix-network
{"nodes":[{"id":"A","label":"User"},{"id":"B","label":"Group"}],"edges":[{"from":"A","to":"B","label":"memberOf"}]}
```

```mermaid
flowchart LR
A --> B
```

```math
x^2 + 1
```

1. Item one
2. Item two
""";
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Chart.Enabled = true;
        options.Network.Enabled = true;
        options.Mermaid.Enabled = true;
        options.Math.Enabled = true;

        string html = MarkdownRendererShell.RenderBodyHtml(markdown, options);
        MarkdownDoc document = html.LoadFromHtml();

        Assert.Collection(document.Blocks,
            block => Assert.IsType<HeadingBlock>(block),
            block => Assert.IsType<HeadingBlock>(block),
            block => Assert.IsType<TableBlock>(block),
            block => {
                var semantic = Assert.IsType<SemanticFencedBlock>(block);
                Assert.Equal("ix-chart", semantic.Language);
                Assert.Equal(MarkdownSemanticKinds.Chart, semantic.SemanticKind);
            },
            block => {
                var semantic = Assert.IsType<SemanticFencedBlock>(block);
                Assert.Equal("ix-network", semantic.Language);
                Assert.Equal(MarkdownSemanticKinds.Network, semantic.SemanticKind);
            },
            block => {
                var semantic = Assert.IsType<SemanticFencedBlock>(block);
                Assert.Equal("mermaid", semantic.Language);
                Assert.Equal(MarkdownSemanticKinds.Mermaid, semantic.SemanticKind);
            },
            block => {
                var semantic = Assert.IsType<SemanticFencedBlock>(block);
                Assert.Equal("math", semantic.Language);
                Assert.Equal(MarkdownSemanticKinds.Math, semantic.SemanticKind);
            },
            block => Assert.IsType<OrderedListBlock>(block));

        string roundTripped = NormalizeMarkdown(document.ToMarkdown());
        Assert.Contains("```ix-chart", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```ix-network", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```mermaid", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```math", roundTripped, StringComparison.Ordinal);
        Assert.Contains("| Check | Expected |", roundTripped, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_ExportedTranscriptVisualPack_WithSemanticRecovery() {
        string markdown = LoadCompatibilityFixture("ix-exported-transcript-visual-pack.md");
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Chart.Enabled = true;
        options.Network.Enabled = true;
        options.Mermaid.Enabled = true;

        string html = MarkdownRendererShell.RenderBodyHtml(markdown, options);
        MarkdownDoc document = html.LoadFromHtml();

        Assert.True(document.Blocks.OfType<HeadingBlock>().Count() >= 8);
        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "Assistant (20:30: 13)");
        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "User (20:36: 21)");
        Assert.True(document.Blocks.OfType<TableBlock>().Count() >= 1);
        Assert.True(document.Blocks.OfType<OrderedListBlock>().Count() >= 1);

        var semanticBlocks = document.Blocks.OfType<SemanticFencedBlock>().ToList();
        Assert.Equal(5, semanticBlocks.Count(block => block.Language == "ix-chart"));
        Assert.Equal(2, semanticBlocks.Count(block => block.Language == "mermaid"));
        Assert.Contains(semanticBlocks, block => block.SemanticKind == MarkdownSemanticKinds.Chart && block.Language == "ix-chart");
        Assert.Contains(semanticBlocks, block => block.SemanticKind == MarkdownSemanticKinds.Mermaid && block.Language == "mermaid");
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"label\": \"Broken\"", StringComparison.Ordinal));

        string roundTripped = NormalizeMarkdown(document.ToMarkdown());
        Assert.Contains("## Proactive checks", roundTripped, StringComparison.Ordinal);
        Assert.Contains("| Check | What to verify | Expected pass signal |", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```ix-chart", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```mermaid", roundTripped, StringComparison.Ordinal);
        Assert.Contains("\"label\": \"Broken\"", roundTripped, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_ExportedTranscriptVisualPack_PreservesSemanticBlockOrder() {
        string markdown = LoadCompatibilityFixture("ix-exported-transcript-visual-pack.md");
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Chart.Enabled = true;
        options.Network.Enabled = true;
        options.Mermaid.Enabled = true;

        string html = MarkdownRendererShell.RenderBodyHtml(markdown, options);
        MarkdownDoc document = html.LoadFromHtml();

        var sequence = document.Blocks
            .Where(block => block is SemanticFencedBlock or TableBlock or OrderedListBlock)
            .Select(static block => block switch {
                SemanticFencedBlock semantic => $"semantic:{semantic.Language}",
                TableBlock => "table",
                OrderedListBlock => "olist",
                _ => "other"
            })
            .ToArray();

        Assert.Equal(
            [
                "table",
                "olist",
                "semantic:ix-chart",
                "semantic:mermaid",
                "semantic:ix-chart",
                "semantic:ix-chart",
                "semantic:ix-chart",
                "semantic:ix-chart",
                "semantic:mermaid"
            ],
            sequence);
    }

    [Fact]
    public void HtmlToMarkdown_RoundTrips_ExportedTranscriptChartSuite_WithSemanticRecovery() {
        string markdown = LoadCompatibilityFixture("ix-exported-transcript-chart-suite.md");
        var options = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        options.Chart.Enabled = true;
        options.Mermaid.Enabled = true;

        string html = MarkdownRendererShell.RenderBodyHtml(markdown, options);
        MarkdownDoc document = html.LoadFromHtml();

        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "Assistant (20:36: 24)");
        Assert.Contains(document.Blocks, block => block is HeadingBlock heading && heading.Text == "Assistant (20:37: 49)");
        Assert.True(document.Blocks.OfType<TableBlock>().Count() >= 1);

        var semanticBlocks = document.Blocks.OfType<SemanticFencedBlock>().ToList();
        Assert.Equal(4, semanticBlocks.Count(block => block.Language == "ix-chart"));
        Assert.Equal(1, semanticBlocks.Count(block => block.Language == "mermaid"));
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"label\": \"Critical\"", StringComparison.Ordinal));
        Assert.Contains(semanticBlocks, block => block.Content.Contains("\"label\": \"Privileged Changes\"", StringComparison.Ordinal));

        var sequence = document.Blocks
            .Where(block => block is SemanticFencedBlock or TableBlock)
            .Select(static block => block switch {
                SemanticFencedBlock semantic => $"semantic:{semantic.Language}",
                TableBlock => "table",
                _ => "other"
            })
            .ToArray();

        Assert.Equal(
            [
                "semantic:ix-chart",
                "semantic:ix-chart",
                "semantic:ix-chart",
                "semantic:ix-chart",
                "table",
                "semantic:mermaid"
            ],
            sequence);

        string roundTripped = NormalizeMarkdown(document.ToMarkdown());
        Assert.Contains("Risk distribution", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```ix-chart", roundTripped, StringComparison.Ordinal);
        Assert.Contains("```mermaid", roundTripped, StringComparison.Ordinal);
        Assert.Contains("Feature | Test block | Expected result | If it fails, likely cause", roundTripped, StringComparison.Ordinal);
    }

    private static string NormalizeMarkdown(string markdown) {
        return markdown.Replace("\r\n", "\n").TrimEnd('\n');
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
}
