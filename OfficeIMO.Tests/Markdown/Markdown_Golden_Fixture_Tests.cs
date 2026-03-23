using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Golden_Fixture_Tests {
    private static readonly Regex MermaidHashAttributeRegex = new(
        "\\sdata-mermaid-hash=\"[^\"]*\"",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex CompilerGeneratedLambdaRegex = new(
        ">b__\\d+_(\\d+)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    [Fact]
    public void MarkdownGolden_ProfileBoundary() {
        string markdown = LoadCompatibilityFixture("portable-profile-boundary.md");
        var htmlOptions = CreatePlainHtmlOptions();

        var officeDoc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateOfficeIMOProfile());
        var portableDoc = MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreatePortableProfile());

        var sb = new StringBuilder();
        AppendSection(sb, "office.ast", BuildDocumentSummary(officeDoc));
        AppendSection(sb, "office.html", NormalizeHtml(officeDoc.ToHtmlFragment(htmlOptions)));
        AppendSection(sb, "portable.ast", BuildDocumentSummary(portableDoc));
        AppendSection(sb, "portable.html", NormalizeHtml(portableDoc.ToHtmlFragment(htmlOptions)));

        AssertGolden("profile-boundary", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxVisuals() {
        string markdown = LoadCompatibilityFixture("ix-visuals.md");

        var generic = MarkdownRendererPresets.CreateStrictMinimal();
        generic.Chart.Enabled = true;
        generic.Network.Enabled = true;

        var ix = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        ix.Chart.Enabled = true;
        ix.Network.Enabled = true;

        var sb = new StringBuilder();
        AppendSection(sb, "generic.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, generic)));
        AppendSection(sb, "ix.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, ix)));

        AssertGolden("ix-visuals", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxTranscriptNormalization() {
        string markdown = LoadCompatibilityFixture("ix-transcript.md");

        var strict = MarkdownRendererPresets.CreateStrictMinimal();
        var chat = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();

        var sb = new StringBuilder();
        AppendSection(sb, "strict.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strict)));
        AppendSection(sb, "chat.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chat)));

        AssertGolden("ix-transcript", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxSourceDerivedSignalFlowNormalization() {
        string markdown = LoadCompatibilityFixture("ix-source-derived-signal-flow.md");

        var strict = MarkdownRendererPresets.CreateStrictMinimal();
        var chat = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();

        var sb = new StringBuilder();
        AppendSection(sb, "strict.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strict)));
        AppendSection(sb, "chat.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chat)));

        AssertGolden("ix-source-derived-signal-flow", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxSourceDerivedHistoricalReplicationArtifacts() {
        string markdown = LoadCompatibilityFixture("ix-source-derived-historical-replication.md");

        var strict = MarkdownRendererPresets.CreateStrictMinimal();
        var chat = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();

        var sb = new StringBuilder();
        AppendSection(sb, "strict.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strict)));
        AppendSection(sb, "chat.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chat)));

        AssertGolden("ix-source-derived-historical-replication", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxSourceDerivedCollapsedMetrics() {
        string markdown = LoadCompatibilityFixture("ix-source-derived-collapsed-metrics.md");

        var strict = MarkdownRendererPresets.CreateStrictMinimal();
        var chat = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();

        var sb = new StringBuilder();
        AppendSection(sb, "strict.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strict)));
        AppendSection(sb, "chat.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chat)));

        AssertGolden("ix-source-derived-collapsed-metrics", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxSourceDerivedHostLabelBulletArtifacts() {
        string markdown = LoadCompatibilityFixture("ix-source-derived-host-label-bullets.md");

        var strict = MarkdownRendererPresets.CreateStrictMinimal();
        var chat = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();

        var sb = new StringBuilder();
        AppendSection(sb, "strict.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strict)));
        AppendSection(sb, "chat.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chat)));

        AssertGolden("ix-source-derived-host-label-bullets", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxSourceDerivedLegacyToolHeadingArtifacts() {
        string markdown = LoadCompatibilityFixture("ix-source-derived-legacy-tool-heading.md");

        var strict = MarkdownRendererPresets.CreateStrictMinimal();
        var chat = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();

        var sb = new StringBuilder();
        AppendSection(sb, "strict.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strict)));
        AppendSection(sb, "chat.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chat)));

        AssertGolden("ix-source-derived-legacy-tool-heading", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxSourceDerivedBrokenResultLeadIn() {
        string markdown = LoadCompatibilityFixture("ix-source-derived-broken-result.md");

        var strict = MarkdownRendererPresets.CreateStrictMinimal();
        var chat = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();

        var sb = new StringBuilder();
        AppendSection(sb, "strict.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strict)));
        AppendSection(sb, "chat.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chat)));

        AssertGolden("ix-source-derived-broken-result", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxSourceDerivedCachedEvidenceNetworkArtifacts() {
        string markdown = LoadCompatibilityFixture("ix-source-derived-cached-evidence-network.md");

        var strict = MarkdownRendererPresets.CreateStrict();
        strict.Network.Enabled = true;

        var chat = MarkdownRendererPresets.CreateIntelligenceXTranscript();
        chat.Network.Enabled = true;

        var sb = new StringBuilder();
        AppendSection(sb, "strict.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strict)));
        AppendSection(sb, "chat.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chat)));

        AssertGolden("ix-source-derived-cached-evidence-network", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxSourceDerivedCachedEvidenceVisualArtifacts() {
        string markdown = LoadCompatibilityFixture("ix-source-derived-cached-evidence-visuals.md");

        var strict = MarkdownRendererPresets.CreateStrict();
        strict.Chart.Enabled = true;

        var chat = MarkdownRendererPresets.CreateIntelligenceXTranscript();
        chat.Chart.Enabled = true;

        var sb = new StringBuilder();
        AppendSection(sb, "strict.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, strict)));
        AppendSection(sb, "chat.html", NormalizeHtml(OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, chat)));

        AssertGolden("ix-source-derived-cached-evidence-visuals", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxPortableExportLegacyJsonVisuals() {
        string markdown = LoadCompatibilityFixture("ix-portable-export-legacy-json-visuals.md");

        var sb = new StringBuilder();
        AppendSection(sb, "ix-compat.markdown", NormalizeText(
            MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForExport(markdown)));
        AppendSection(sb, "portable-export.markdown", NormalizeText(
            MarkdownTranscriptPreparation.PrepareIntelligenceXTranscriptForExport(
                markdown,
                MarkdownVisualFenceLanguageMode.GenericSemanticFence)));

        AssertGolden("ix-portable-export-legacy-json-visuals", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_HtmlRichAst() {
        string html = LoadCompatibilityFixture("html-rich-ast.html");

        MarkdownDoc document = html.LoadFromHtml();
        string officeMarkdown = document.ToMarkdown(MarkdownWriteOptions.CreateOfficeIMOProfile());
        string portableMarkdown = html.ToMarkdown(HtmlToMarkdownOptions.CreatePortableProfile());
        string renderedHtml = document.ToHtmlFragment(CreatePlainHtmlOptions());

        var sb = new StringBuilder();
        AppendSection(sb, "ast", BuildDocumentSummary(document));
        AppendSection(sb, "office.markdown", NormalizeText(officeMarkdown));
        AppendSection(sb, "portable.markdown", NormalizeText(portableMarkdown));
        AppendSection(sb, "rendered.html", NormalizeHtml(renderedHtml));

        AssertGolden("html-rich-ast", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_HtmlSharedVisualHosts() {
        string html = LoadHtmlFixture("shared-visual-hosts.html");

        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/visuals/archive.html")
        });
        string officeMarkdown = document.ToMarkdown(MarkdownWriteOptions.CreateOfficeIMOProfile());
        string portableMarkdown = html.ToMarkdown(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/visuals/archive.html"),
            MarkdownWriteOptions = MarkdownWriteOptions.CreatePortableProfile()
        });
        string renderedHtml = document.ToHtmlFragment(CreatePlainHtmlOptions());

        var sb = new StringBuilder();
        AppendSection(sb, "ast", BuildDocumentSummary(document));
        AppendSection(sb, "office.markdown", NormalizeText(officeMarkdown));
        AppendSection(sb, "portable.markdown", NormalizeText(portableMarkdown));
        AppendSection(sb, "rendered.html", NormalizeHtml(renderedHtml));

        AssertGolden("html-shared-visual-hosts", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_HtmlPublisherFigureFixtures() {
        var sb = new StringBuilder();
        AppendHtmlFixtureSnapshot(sb, "linked", "publisher-linked-picture-article.html");
        AppendHtmlFixtureSnapshot(sb, "noscript", "publisher-noscript-linked-picture-article.html");
        AppendHtmlFixtureSnapshot(sb, "art-direction", "publisher-art-direction-picture-article.html");
        AppendHtmlFixtureSnapshot(sb, "width-descriptor", "publisher-width-descriptor-picture-article.html");

        AssertGolden("html-publisher-figure-fixtures", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxCompatibilityTranscriptVisualPackRoundTrip() {
        string markdown = LoadCompatibilityFixture("ix-compat-transcript-visual-pack.md");
        var ix = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        ix.Chart.Enabled = true;
        ix.Network.Enabled = true;
        ix.Mermaid.Enabled = true;

        string html = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, ix);
        MarkdownDoc document = html.LoadFromHtml();

        var sb = new StringBuilder();
        AppendSection(sb, "ix.html", NormalizeHtml(html));
        AppendSection(sb, "recovered.ast", BuildDocumentSummary(document));
        AppendSection(sb, "roundtrip.markdown", NormalizeText(document.ToMarkdown()));

        AssertGolden("ix-compat-transcript-visual-pack", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxCompatibilityTranscriptChartSuiteRoundTrip() {
        string markdown = LoadCompatibilityFixture("ix-compat-transcript-chart-suite.md");
        var ix = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        ix.Chart.Enabled = true;
        ix.Mermaid.Enabled = true;

        string html = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, ix);
        MarkdownDoc document = html.LoadFromHtml();

        var sb = new StringBuilder();
        AppendSection(sb, "ix.html", NormalizeHtml(html));
        AppendSection(sb, "recovered.ast", BuildDocumentSummary(document));
        AppendSection(sb, "roundtrip.markdown", NormalizeText(document.ToMarkdown()));

        AssertGolden("ix-compat-transcript-chart-suite", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_IxCompatibilityPipeline() {
        string legacyToolHeadingMarkdown = LoadCompatibilityFixture("ix-source-derived-legacy-tool-heading.md");
        string cachedEvidenceVisualMarkdown = LoadCompatibilityFixture("ix-source-derived-cached-evidence-visuals.md");

        var legacyOptions = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        var cachedEvidenceOptions = MarkdownRendererPresets.CreateIntelligenceXTranscript();
        cachedEvidenceOptions.Chart.Enabled = true;

        var legacyResult = OfficeIMO.MarkdownRenderer.MarkdownRenderer.ParseDocumentResult(legacyToolHeadingMarkdown, legacyOptions);
        var cachedEvidenceResult = OfficeIMO.MarkdownRenderer.MarkdownRenderer.ParseDocumentResult(cachedEvidenceVisualMarkdown, cachedEvidenceOptions);

        var sb = new StringBuilder();
        AppendRendererParseResultSummary(sb, "legacy-tool-heading", legacyResult);
        AppendRendererParseResultSummary(sb, "cached-evidence-visuals", cachedEvidenceResult);

        AssertGolden("ix-compatibility-pipeline", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_RendererCombinedPipeline() {
        var options = new MarkdownRendererOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };
        options.MarkdownPreProcessors.Add((markdown, _) =>
            markdown.Replace("```mermaid\nflowchart LR", "```mermaid\ngraph TD"));
        options.ReaderOptions.DocumentTransforms.Add(new GoldenReaderAppendParagraphTransform("reader tail"));
        options.DocumentTransforms.Add(new GoldenRendererAppendParagraphTransform("renderer tail"));

        var result = OfficeIMO.MarkdownRenderer.MarkdownRenderer.ParseDocumentResult(
            "```mermaidflowchart LR A-->B\n```",
            options);

        var sb = new StringBuilder();
        AppendRendererParseResultSummary(sb, "renderer-combined", result);

        AssertGolden("renderer-combined-pipeline", sb.ToString().TrimEnd());
    }

    [Fact]
    public void MarkdownGolden_RendererSourceAwarePipeline() {
        var options = new MarkdownRendererOptions();
        options.ReaderOptions.DocumentTransforms.Add(new MarkdownCompactHeadingBoundaryTransform());
        options.DocumentTransforms.Add(new GoldenRendererAppendParagraphTransform("renderer tail"));

        var result = OfficeIMO.MarkdownRenderer.MarkdownRenderer.ParseDocumentResult(
            "previous shutdown was unexpected### Reason",
            options);

        var sb = new StringBuilder();
        AppendRendererParseResultSummary(sb, "renderer-source-aware", result);

        AssertGolden("renderer-source-aware-pipeline", sb.ToString().TrimEnd());
    }

    private static HtmlOptions CreatePlainHtmlOptions() {
        return new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        };
    }

    private static string BuildDocumentSummary(MarkdownDoc document) {
        var sb = new StringBuilder();
        AppendBlockList(sb, document.Blocks, 0);
        return sb.ToString().TrimEnd();
    }

    private static void AppendBlockList(StringBuilder sb, IReadOnlyList<IMarkdownBlock> blocks, int indent) {
        for (int i = 0; i < blocks.Count; i++) {
            AppendBlock(sb, blocks[i], indent, i);
        }
    }

    private static void AppendBlock(StringBuilder sb, IMarkdownBlock block, int indent, int index) {
        string prefix = new string(' ', indent * 2);
        sb.Append(prefix)
            .Append(index.ToString(CultureInfo.InvariantCulture))
            .Append(": ")
            .AppendLine(DescribeBlock(block));

        switch (block) {
            case CalloutBlock callout:
                AppendBlockList(sb, callout.ChildBlocks, indent + 1);
                break;
            case QuoteBlock quote:
                AppendBlockList(sb, quote.ChildBlocks, indent + 1);
                break;
            case UnorderedListBlock unordered:
                AppendListItems(sb, unordered.Items, indent + 1);
                break;
            case OrderedListBlock ordered:
                AppendListItems(sb, ordered.Items, indent + 1);
                break;
            case DefinitionListBlock definitionList:
                AppendDefinitionEntries(sb, definitionList, indent + 1);
                break;
            case FootnoteDefinitionBlock footnote:
                AppendParagraphBlocks(sb, footnote.ParagraphBlocks, indent + 1);
                break;
            case DetailsBlock details:
                if (details.Summary != null) {
                    sb.Append(new string(' ', (indent + 1) * 2))
                        .Append("summary: ")
                        .AppendLine(EscapeSingleLine(details.Summary.Inlines.RenderMarkdown()));
                }
                AppendBlockList(sb, details.ChildBlocks, indent + 1);
                break;
            case TableBlock table:
                AppendTableSummary(sb, table, indent + 1);
                break;
        }
    }

    private static void AppendListItems(StringBuilder sb, IReadOnlyList<ListItem> items, int indent) {
        string prefix = new string(' ', indent * 2);
        for (int i = 0; i < items.Count; i++) {
            var item = items[i];
            sb.Append(prefix)
                .Append("item[")
                .Append(i.ToString(CultureInfo.InvariantCulture))
                .Append("]: task=")
                .Append(item.IsTask ? (item.Checked ? "checked" : "unchecked") : "no")
                .Append(" content=\"")
                .Append(EscapeSingleLine(item.Content.RenderMarkdown()))
                .AppendLine("\"");

            AppendBlockList(sb, item.Children, indent + 1);
        }
    }

    private static void AppendDefinitionEntries(StringBuilder sb, DefinitionListBlock definitionList, int indent) {
        string prefix = new string(' ', indent * 2);
        for (int i = 0; i < definitionList.Entries.Count; i++) {
            var entry = definitionList.Entries[i];
            sb.Append(prefix)
                .Append("entry[")
                .Append(i.ToString(CultureInfo.InvariantCulture))
                .Append("]: term=\"")
                .Append(EscapeSingleLine(entry.TermMarkdown))
                .AppendLine("\"");

            AppendBlockList(sb, entry.DefinitionBlocks, indent + 1);
        }
    }

    private static void AppendParagraphBlocks(StringBuilder sb, IReadOnlyList<ParagraphBlock> paragraphs, int indent) {
        string prefix = new string(' ', indent * 2);
        for (int i = 0; i < paragraphs.Count; i++) {
            sb.Append(prefix)
                .Append("paragraph[")
                .Append(i.ToString(CultureInfo.InvariantCulture))
                .Append("]: \"")
                .Append(EscapeSingleLine(paragraphs[i].Inlines.RenderMarkdown()))
                .AppendLine("\"");
        }
    }

    private static void AppendTableSummary(StringBuilder sb, TableBlock table, int indent) {
        string prefix = new string(' ', indent * 2);
        if (table.HeaderCells.Count > 0) {
            sb.Append(prefix)
                .Append("headers: ")
                .AppendLine(string.Join(" || ", table.HeaderCells.Select(cell => "\"" + EscapeSingleLine(cell.Markdown) + "\"")));
        }

        for (int rowIndex = 0; rowIndex < table.RowCells.Count; rowIndex++) {
            sb.Append(prefix)
                .Append("row[")
                .Append(rowIndex.ToString(CultureInfo.InvariantCulture))
                .Append("]: ")
                .AppendLine(string.Join(" || ", table.RowCells[rowIndex].Select(cell => "\"" + EscapeSingleLine(cell.Markdown) + "\"")));
        }
    }

    private static string DescribeBlock(IMarkdownBlock block) {
        return block switch {
            HeadingBlock heading => $"Heading(level={heading.Level}, text=\"{EscapeSingleLine(heading.Text)}\")",
            ParagraphBlock paragraph => $"Paragraph(\"{EscapeSingleLine(paragraph.Inlines.RenderMarkdown())}\")",
            CalloutBlock callout => $"Callout(kind={callout.Kind}, title=\"{EscapeSingleLine(callout.TitleInlines.RenderMarkdown())}\")",
            QuoteBlock => "Quote",
            UnorderedListBlock unordered => $"UnorderedList(items={unordered.Items.Count.ToString(CultureInfo.InvariantCulture)})",
            OrderedListBlock ordered => $"OrderedList(start={ordered.Start.ToString(CultureInfo.InvariantCulture)}, items={ordered.Items.Count.ToString(CultureInfo.InvariantCulture)})",
            TableBlock table => $"Table(headers={table.HeaderCells.Count.ToString(CultureInfo.InvariantCulture)}, rows={table.RowCells.Count.ToString(CultureInfo.InvariantCulture)})",
            DefinitionListBlock definitionList => $"DefinitionList(entries={definitionList.Entries.Count.ToString(CultureInfo.InvariantCulture)})",
            FootnoteDefinitionBlock footnote => $"Footnote(label={footnote.Label})",
            DetailsBlock details => $"Details(open={details.Open.ToString().ToLowerInvariant()})",
            SemanticFencedBlock semantic => $"Semantic(kind={semantic.SemanticKind}, language={semantic.Language}, text=\"{EscapeSingleLine(AbbreviateSemanticContent(semantic.Content))}\")",
            _ => block.GetType().Name
        };
    }

    private static void AppendSection(StringBuilder sb, string name, string content) {
        sb.Append('[').Append(name).AppendLine("]");
        sb.AppendLine(content);
        sb.AppendLine();
    }

    private static void AppendRendererParseResultSummary(
        StringBuilder sb,
        string name,
        MarkdownRendererParseResult result) {
        AppendSection(sb, name + ".preprocessed", NormalizeText(result.PreprocessedMarkdown));
        AppendSection(sb, name + ".preprocessors", BuildPreProcessorDiagnosticSummary(result.PreProcessorDiagnostics));
        AppendSection(sb, name + ".transforms", BuildTransformDiagnosticSummary(result.TransformDiagnostics));
        AppendSection(sb, name + ".ast", BuildDocumentSummary(result.Document));
    }

    private static void AppendHtmlFixtureSnapshot(StringBuilder sb, string name, string fixtureName) {
        string html = LoadHtmlFixture(fixtureName);
        MarkdownDoc document = html.LoadFromHtml(new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/world/live/storm-update.html")
        });

        AppendSection(sb, name + ".ast", BuildDocumentSummary(document));
        AppendSection(sb, name + ".markdown", NormalizeText(document.ToMarkdown()));
        AppendSection(sb, name + ".html", NormalizeInlineAwareHtml(document.ToHtmlFragment(CreatePlainHtmlOptions())));
    }

    private static string BuildPreProcessorDiagnosticSummary(
        IReadOnlyList<MarkdownRendererPreProcessorDiagnostic> diagnostics) {
        if (diagnostics == null || diagnostics.Count == 0) {
            return "(none)";
        }

        var sb = new StringBuilder();
        for (var i = 0; i < diagnostics.Count; i++) {
            var diagnostic = diagnostics[i];
            sb.Append(diagnostic.Stage)
                .Append(" changed=")
                .Append(diagnostic.Changed ? "true" : "false")
                .Append(" before=")
                .Append(diagnostic.LengthBefore.ToString(CultureInfo.InvariantCulture))
                .Append(" after=")
                .Append(diagnostic.LengthAfter.ToString(CultureInfo.InvariantCulture));

            if (!string.IsNullOrWhiteSpace(diagnostic.ProcessorName)) {
                sb.Append(" processor=")
                    .Append(GetSimpleTypeName(diagnostic.ProcessorName));
            }

            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    private static string BuildTransformDiagnosticSummary(
        IReadOnlyList<MarkdownDocumentTransformDiagnostic> diagnostics) {
        if (diagnostics == null || diagnostics.Count == 0) {
            return "(none)";
        }

        var sb = new StringBuilder();
        for (var i = 0; i < diagnostics.Count; i++) {
            var diagnostic = diagnostics[i];
            sb.Append(diagnostic.Source)
                .Append(" ")
                .Append(GetSimpleTypeName(diagnostic.TransformName))
                .Append(" before=")
                .Append(diagnostic.BlockCountBefore.ToString(CultureInfo.InvariantCulture))
                .Append(" after=")
                .Append(diagnostic.BlockCountAfter.ToString(CultureInfo.InvariantCulture))
                .Append(" changed-before=")
                .Append(diagnostic.ChangedBlockStartBefore.ToString(CultureInfo.InvariantCulture))
                .Append("+")
                .Append(diagnostic.ChangedBlockCountBefore.ToString(CultureInfo.InvariantCulture))
                .Append(" changed-after=")
                .Append(diagnostic.ChangedBlockStartAfter.ToString(CultureInfo.InvariantCulture))
                .Append("+")
                .Append(diagnostic.ChangedBlockCountAfter.ToString(CultureInfo.InvariantCulture))
                .Append(" span=")
                .Append(FormatSourceSpan(diagnostic.AffectedSourceSpan))
                .Append(" replaced=")
                .Append(diagnostic.ReplacedDocument ? "true" : "false")
                .AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    private static string GetSimpleTypeName(string value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        var lastDot = value.LastIndexOf('.');
        return lastDot >= 0 && lastDot + 1 < value.Length
            ? value.Substring(lastDot + 1)
            : value;
    }

    private static string FormatSourceSpan(MarkdownSourceSpan? span) {
        if (!span.HasValue) {
            return "-";
        }

        return span.Value.ToString();
    }

    private static string NormalizeHtml(string html) {
        if (string.IsNullOrWhiteSpace(html)) {
            return string.Empty;
        }

        html = MermaidHashAttributeRegex.Replace(html, " data-mermaid-hash=\"{normalized}\"");

        var sb = new StringBuilder(html.Length);
        bool inTag = false;
        bool lastWasWhitespace = false;

        for (int i = 0; i < html.Length; i++) {
            char ch = html[i];
            if (ch == '<') {
                if (!inTag && lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                    sb.Append(' ');
                }

                inTag = true;
                lastWasWhitespace = false;
                sb.Append(ch);
                continue;
            }

            if (ch == '>') {
                inTag = false;
                lastWasWhitespace = false;
                sb.Append(ch);
                continue;
            }

            if (inTag) {
                sb.Append(ch);
                continue;
            }

            if (char.IsWhiteSpace(ch)) {
                lastWasWhitespace = true;
                continue;
            }

            if (lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                sb.Append(' ');
            }

            lastWasWhitespace = false;
            sb.Append(ch);
        }

        return sb.ToString()
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim();
    }

    private static string NormalizeInlineAwareHtml(string html) {
        if (string.IsNullOrWhiteSpace(html)) {
            return string.Empty;
        }

        html = MermaidHashAttributeRegex.Replace(html, " data-mermaid-hash=\"{normalized}\"");

        var sb = new StringBuilder(html.Length);
        bool inTag = false;
        bool lastWasWhitespace = false;

        for (int i = 0; i < html.Length; i++) {
            char ch = html[i];
            if (ch == '<') {
                if (!inTag && lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                    sb.Append(' ');
                }

                inTag = true;
                lastWasWhitespace = false;
                sb.Append(ch);
                continue;
            }

            if (ch == '>') {
                inTag = false;
                lastWasWhitespace = false;
                sb.Append(ch);
                continue;
            }

            if (inTag) {
                sb.Append(ch);
                continue;
            }

            if (char.IsWhiteSpace(ch)) {
                lastWasWhitespace = true;
                continue;
            }

            if (lastWasWhitespace && sb.Length > 0) {
                sb.Append(' ');
            }

            lastWasWhitespace = false;
            sb.Append(ch);
        }

        return sb.ToString()
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim();
    }

    private static string NormalizeText(string value) {
        return CompilerGeneratedLambdaRegex.Replace(
                MermaidHashAttributeRegex.Replace(value, " data-mermaid-hash=\"{normalized}\""),
                ">b__{normalized}_$1")
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim();
    }

    private static string NormalizeFixtureText(string value) {
        return (value ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
    }

    private static string EscapeSingleLine(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        return value!
            .Replace("\r\n", "\\n")
            .Replace('\r', '\n')
            .Replace("\n", "\\n");
    }

    private static string AbbreviateSemanticContent(string? value) {
        const int maxLength = 80;
        if (string.IsNullOrEmpty(value) || value!.Length <= maxLength) {
            return value ?? string.Empty;
        }

        return value.Substring(0, maxLength - 3) + "...";
    }

    private static void AssertGolden(string name, string actualSnapshot) {
        string expectedPath = GetExpectedPath(name);
        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_GOLDEN"), "1", StringComparison.Ordinal)) {
            Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
            File.WriteAllText(expectedPath, actualSnapshot + Environment.NewLine, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            return;
        }

        if (!File.Exists(expectedPath)) {
            throw new FileNotFoundException(
                "Golden snapshot missing. Set OFFICEIMO_UPDATE_GOLDEN=1 and re-run this test to generate it.",
                expectedPath);
        }

        string expected = File.ReadAllText(expectedPath, Encoding.UTF8);
        Assert.Equal(NormalizeText(expected), NormalizeText(actualSnapshot));
    }

    private static string LoadCompatibilityFixture(string name) {
        return NormalizeFixtureText(File.ReadAllText(Path.Combine(GetTestsProjectRoot(), "Markdown", "Fixtures", "Compatibility", name)));
    }

    private static string LoadHtmlFixture(string name) {
        return NormalizeFixtureText(File.ReadAllText(Path.Combine(GetTestsProjectRoot(), "Markdown", "Fixtures", name)));
    }

    private static string GetExpectedPath(string name) {
        return Path.Combine(GetTestsProjectRoot(), "Markdown", "Golden", "Expected", name + ".snapshot.txt");
    }

    private static string GetTestsProjectRoot() {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null) {
            string candidate = Path.Combine(dir.FullName, "OfficeIMO.Tests.csproj");
            if (File.Exists(candidate)) {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate OfficeIMO.Tests project root from test runtime base directory.");
    }

    private sealed class GoldenRendererAppendParagraphTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            document.Add(new ParagraphBlock(new InlineSequence().Text(text)));
            return document;
        }
    }

    private sealed class GoldenReaderAppendParagraphTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            document.Add(new ParagraphBlock(new InlineSequence().Text(text)));
            return document;
        }
    }
}

