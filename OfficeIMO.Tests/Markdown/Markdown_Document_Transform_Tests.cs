using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Document_Transform_Tests {
    [Fact]
    public void MarkdownReader_Applies_DocumentTransforms_In_Order() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new AppendParagraphTransform("first"));
        options.DocumentTransforms.Add(new AppendParagraphTransform("second"));

        var document = MarkdownReader.Parse("Base paragraph.", options);

        Assert.Equal(
            NormalizeMarkdown("""
Base paragraph.

first

second
"""),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void MarkdownDocumentTransformPipeline_Collects_Diagnostics() {
        var diagnostics = new List<MarkdownDocumentTransformDiagnostic>();
        var document = MarkdownDoc.Create();
        document.Add(new ParagraphBlock(new InlineSequence().Text("Base")));
        var transforms = new IMarkdownDocumentTransform[] {
            new AppendParagraphTransform("first"),
            new AppendParagraphTransform("second")
        };

        var transformed = MarkdownDocumentTransformPipeline.Apply(
            document,
            transforms,
            new MarkdownDocumentTransformContext(
                MarkdownDocumentTransformSource.MarkdownReader,
                MarkdownReaderOptions.CreateOfficeIMOProfile(),
                sourceOptions: null,
                diagnostics));

        Assert.Equal(2, diagnostics.Count);
        Assert.All(diagnostics, diagnostic => Assert.Equal(MarkdownDocumentTransformSource.MarkdownReader, diagnostic.Source));
        Assert.Contains(nameof(AppendParagraphTransform), diagnostics[0].TransformName, StringComparison.Ordinal);
        Assert.Equal(1, diagnostics[0].BlockCountBefore);
        Assert.Equal(2, diagnostics[0].BlockCountAfter);
        Assert.False(diagnostics[0].ReplacedDocument);
        Assert.Equal(1, diagnostics[0].ChangedBlockStartBefore);
        Assert.Equal(0, diagnostics[0].ChangedBlockCountBefore);
        Assert.Equal(1, diagnostics[0].ChangedBlockStartAfter);
        Assert.Equal(1, diagnostics[0].ChangedBlockCountAfter);
        Assert.Null(diagnostics[0].AffectedSourceSpan);
        Assert.Equal(3, transformed.Blocks.Count);
    }

    [Fact]
    public void MarkdownDocumentTransformPipeline_Collects_AffectedSourceSpan_When_SyntaxTree_Is_Available() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        var parseResult = MarkdownReader.ParseWithSyntaxTree("previous shutdown was unexpected### Reason", options);
        var diagnostics = new List<MarkdownDocumentTransformDiagnostic>();
        var transforms = new IMarkdownDocumentTransform[] {
            new MarkdownCompactHeadingBoundaryTransform()
        };

        var transformed = MarkdownDocumentTransformPipeline.Apply(
            parseResult.Document,
            transforms,
            new MarkdownDocumentTransformContext(
                MarkdownDocumentTransformSource.MarkdownReader,
                options,
                sourceOptions: null,
                diagnostics,
                parseResult.SyntaxTree));

        var diagnostic = Assert.Single(diagnostics);
        Assert.Equal(0, diagnostic.ChangedBlockStartBefore);
        Assert.Equal(1, diagnostic.ChangedBlockCountBefore);
        Assert.Equal(0, diagnostic.ChangedBlockStartAfter);
        Assert.Equal(2, diagnostic.ChangedBlockCountAfter);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 42), diagnostic.AffectedSourceSpan);
        Assert.Equal("Document > Paragraph", diagnostic.AffectedOriginalBlockPath);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 42), diagnostic.AffectedOriginalBlockSpan);
        Assert.Equal(2, transformed.Blocks.Count);
    }

    [Fact]
    public void MarkdownReader_ParseWithSyntaxTreeAndDiagnostics_Collects_TransformDiagnostics_InOneCall() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownCompactHeadingBoundaryTransform());

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("previous shutdown was unexpected### Reason", options);

        Assert.Equal(2, result.Document.Blocks.Count);
        Assert.Single(result.TransformDiagnostics);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 42), result.TransformDiagnostics[0].AffectedSourceSpan);
    }

    [Fact]
    public void MarkdownJsonVisualCodeBlockTransform_Upgrades_LegacyJsonCodeBlock_To_SemanticBlock() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(
            new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence));

        var document = MarkdownReader.Parse("""
```json
{"type":"bar","data":{"labels":["A"],"datasets":[{"label":"Count","data":[1]}]}}
```
""", options);

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("ix-chart", block.Language);
    }

    [Fact]
    public void MarkdownJsonVisualCodeBlockTransform_Is_Idempotent() {
        var transform = new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.GenericSemanticFence);
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(transform);

        var document = MarkdownReader.Parse("""
```json
{"nodes":[{"id":"A","label":"Root"}],"edges":[{"from":"A","to":"B"}]}
```
""", options);

        var once = NormalizeMarkdown(document.ToMarkdown());
        var twice = NormalizeMarkdown(MarkdownDocumentTransformPipeline.Apply(
            document,
            [transform],
            new MarkdownDocumentTransformContext(MarkdownDocumentTransformSource.MarkdownReader, options)).ToMarkdown());

        Assert.Equal(once, twice);
    }

    [Fact]
    public void MarkdownStandaloneHashHeadingSeparatorTransform_Removes_Empty_Hash_Heading_Before_Real_Heading() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownStandaloneHashHeadingSeparatorTransform());

        var document = MarkdownReader.Parse("""
#

## Result
""", options);

        Assert.Collection(document.Blocks,
            block => {
                var heading = Assert.IsType<HeadingBlock>(block);
                Assert.Equal(2, heading.Level);
                Assert.Equal("Result", heading.Text);
            });
    }

    [Fact]
    public void MarkdownStandaloneHashHeadingSeparatorTransform_Preserves_Ordinary_Empty_Hash_Heading_When_Not_Followed_By_Heading() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownStandaloneHashHeadingSeparatorTransform());

        var document = MarkdownReader.Parse("""
#

Paragraph body.
""", options);

        Assert.Collection(document.Blocks,
            block => {
                var heading = Assert.IsType<HeadingBlock>(block);
                Assert.Equal(1, heading.Level);
                Assert.True(string.IsNullOrWhiteSpace(heading.Text));
            },
            block => Assert.IsType<ParagraphBlock>(block));
    }

    [Fact]
    public void MarkdownCompactHeadingBoundaryTransform_Splits_Compact_Paragraph_Heading_Boundaries() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownCompactHeadingBoundaryTransform());

        var document = MarkdownReader.Parse("previous shutdown was unexpected### Reason", options);

        Assert.Collection(document.Blocks,
            block => {
                var paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Equal("previous shutdown was unexpected", paragraph.Inlines.RenderMarkdown());
            },
            block => {
                var heading = Assert.IsType<HeadingBlock>(block);
                Assert.Equal(3, heading.Level);
                Assert.Equal("Reason", heading.Text);
            });
    }

    [Fact]
    public void MarkdownCompactHeadingBoundaryTransform_DoesNotSplit_CodeSpan_Content() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownCompactHeadingBoundaryTransform());

        var document = MarkdownReader.Parse("Use `unexpected### Reason` as captured text.", options);

        Assert.Collection(document.Blocks,
            block => {
                var paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Equal("Use `unexpected### Reason` as captured text.", paragraph.Inlines.RenderMarkdown());
            });
    }

    [Fact]
    public void MarkdownColonListBoundaryTransform_Splits_Paragraph_List_Boundaries() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownColonListBoundaryTransform());

        var document = MarkdownReader.Parse("Następny najlepszy krok:- **`ad_domain_controller_facts`**", options);

        Assert.Collection(document.Blocks,
            block => {
                var paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Equal("Następny najlepszy krok:", paragraph.Inlines.RenderMarkdown());
            },
            block => {
                var list = Assert.IsType<UnorderedListBlock>(block);
                var item = Assert.Single(list.Items);
                Assert.Equal("**`ad_domain_controller_facts`**", item.Content.RenderMarkdown());
            });
    }

    [Fact]
    public void MarkdownColonListBoundaryTransform_DoesNotSplit_CodeSpan_Content() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownColonListBoundaryTransform());

        var document = MarkdownReader.Parse("Use `Next step:- **Item**` as captured text.", options);

        Assert.Collection(document.Blocks,
            block => {
                var paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Equal("Use `Next step:- **Item**` as captured text.", paragraph.Inlines.RenderMarkdown());
            });
    }

    [Fact]
    public void MarkdownHeadingListBoundaryTransform_Splits_Heading_List_Boundaries() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownHeadingListBoundaryTransform());

        var document = MarkdownReader.Parse("## Summary- **Replication:** healthy", options);

        Assert.Collection(document.Blocks,
            block => {
                var heading = Assert.IsType<HeadingBlock>(block);
                Assert.Equal(2, heading.Level);
                Assert.Equal("Summary", heading.Text);
            },
            block => {
                var list = Assert.IsType<UnorderedListBlock>(block);
                var item = Assert.Single(list.Items);
                Assert.Equal("**Replication:** healthy", item.Content.RenderMarkdown());
            });
    }

    [Fact]
    public void MarkdownHeadingListBoundaryTransform_DoesNotSplit_CodeSpan_Content() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownHeadingListBoundaryTransform());

        var document = MarkdownReader.Parse("## Summary `- **Replication:**` tail", options);

        Assert.Collection(document.Blocks,
            block => {
                var heading = Assert.IsType<HeadingBlock>(block);
                Assert.Equal("Summary - **Replication:** tail", heading.Text);
            });
    }

    [Fact]
    public void MarkdownCompactStrongLabelListBoundaryTransform_Splits_List_Items() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownCompactStrongLabelListBoundaryTransform());

        var document = MarkdownReader.Parse("- **Replication:** wcześniej zdrowa ✅- **FSMO:** technicznie OK", options);

        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        Assert.Equal(2, list.Items.Count);
        Assert.Equal("**Replication:** wcześniej zdrowa ✅", list.Items[0].Content.RenderMarkdown());
        Assert.Equal("**FSMO:** technicznie OK", list.Items[1].Content.RenderMarkdown());
    }

    [Fact]
    public void MarkdownCompactStrongLabelListBoundaryTransform_DoesNotSplit_CodeSpan_Content() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownCompactStrongLabelListBoundaryTransform());

        var document = MarkdownReader.Parse("✅`- **FSMO:**` tail", options);

        Assert.Collection(document.Blocks,
            block => {
                var paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Equal("✅`- **FSMO:**` tail", paragraph.Inlines.RenderMarkdown());
            });
    }

    [Fact]
    public void MarkdownListParagraphStrongArtifactTransform_Repairs_Malformed_List_Strong_Runs() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownListParagraphStrongArtifactTransform(new MarkdownInputNormalizationOptions {
            NormalizeLooseStrongDelimiters = true,
            NormalizeDanglingTrailingStrongListClosers = true,
            NormalizeMetricValueStrongRuns = true
        }));

        var document = MarkdownReader.Parse("""
- Overall health ****healthy****
- Overall health ✅ Healthy****
- Overall health ******healthy**
- Overall health **✅****Healthy**
- LDAP/LDAPS across all DCs **healthy on FQDN endpoints for all 5 servers*
""", options);

        var markdown = NormalizeMarkdown(document.ToMarkdown());
        Assert.Contains("- Overall health **healthy**", markdown, StringComparison.Ordinal);
        Assert.Contains("- LDAP/LDAPS across all DCs **healthy on FQDN endpoints for all 5 servers**", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("**✅****Healthy**", markdown, StringComparison.Ordinal);
        Assert.Contains("Healthy", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Applies_DocumentTransforms_To_IntermediateDocument() {
        var options = new HtmlToMarkdownOptions();
        options.DocumentTransforms.Add(
            new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.GenericSemanticFence));

        var document = """
<pre><code class="language-json">{"type":"bar","data":{"labels":["A"],"datasets":[{"label":"Count","data":[1]}]}}</code></pre>
""".LoadFromHtml(options);

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(document.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("chart", block.Language);
        Assert.Equal(
            NormalizeMarkdown("""
```chart
{"type":"bar","data":{"labels":["A"],"datasets":[{"label":"Count","data":[1]}]}}
```
"""),
            NormalizeMarkdown(document.ToMarkdown()));
    }

    [Fact]
    public void HtmlToMarkdown_Can_Apply_InlineNormalizationTransform_To_IntermediateDocument() {
        var options = new HtmlToMarkdownOptions();
        options.DocumentTransforms.Add(new MarkdownInlineNormalizationTransform(new MarkdownInputNormalizationOptions {
            NormalizeTightParentheticalSpacing = true,
            NormalizeTightColonSpacing = true
        }));

        var document = """
<p><strong>Deleted object remnants</strong>(SID left in ACL path)</p>
<p>Why it matters:missing evidence</p>
""".LoadFromHtml(options);

        var html = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<strong>Deleted object remnants</strong> (SID left in ACL path)", html, StringComparison.Ordinal);
        Assert.Contains("Why it matters: missing evidence", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Can_Trim_LooseStrongWhitespace_Via_InlineNormalizationTransform() {
        var options = new HtmlToMarkdownOptions();
        options.DocumentTransforms.Add(new MarkdownInlineNormalizationTransform(new MarkdownInputNormalizationOptions {
            NormalizeLooseStrongDelimiters = true
        }));

        var document = """
<p><strong> LDAP/Kerberos health on all DCs </strong> next</p>
""".LoadFromHtml(options);

        var html = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.Contains("<strong>LDAP/Kerberos health on all DCs</strong> next", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_Can_Normalize_TightArrowStrongBoundary_Via_InlineNormalizationTransform() {
        var options = new HtmlToMarkdownOptions();
        options.DocumentTransforms.Add(new MarkdownInlineNormalizationTransform(new MarkdownInputNormalizationOptions {
            NormalizeTightArrowStrongBoundaries = true
        }));

        var document = """
<p>Signal -&gt;<strong>Why it matters:</strong> coverage is thin</p>
""".LoadFromHtml(options);

        Assert.Contains("Signal -> **Why it matters:** coverage is thin", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownSimpleDefinitionListParagraphTransform_Expands_Simple_Definition_List_Entries() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownSimpleDefinitionListParagraphTransform());

        var document = MarkdownReader.Parse("""
Status: healthy
Impact: none
""", options);

        var markdown = NormalizeMarkdown(document.ToMarkdown());
        var html = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Equal(NormalizeMarkdown("""
Status: healthy

Impact: none
"""), markdown);
        Assert.DoesNotContain("<dl>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownSimpleDefinitionListParagraphTransform_Preserves_Complex_Definition_List_Entries() {
        var transform = new MarkdownSimpleDefinitionListParagraphTransform();
        var document = MarkdownDoc.Create();
        var definitions = new DefinitionListBlock();
        definitions.AddEntry(new DefinitionListEntry(
            MarkdownReader.ParseInlineText("Status"),
            new IMarkdownBlock[] { new ParagraphBlock(MarkdownReader.ParseInlineText("healthy")) }));
        definitions.AddEntry(new DefinitionListEntry(
            MarkdownReader.ParseInlineText("Evidence"),
            new IMarkdownBlock[] {
                new ParagraphBlock(MarkdownReader.ParseInlineText("See logs")),
                new QuoteBlock(new[] { "quoted context" })
            }));
        document.Add(definitions);

        var transformed = MarkdownDocumentTransformPipeline.Apply(
            document,
            [transform],
            new MarkdownDocumentTransformContext(MarkdownDocumentTransformSource.MarkdownReader, MarkdownReaderOptions.CreateOfficeIMOProfile()));

        Assert.Collection(transformed.Blocks,
            block => {
                var paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Equal("Status: healthy", paragraph.Inlines.RenderMarkdown());
            },
            block => {
                var remaining = Assert.IsType<DefinitionListBlock>(block);
                var entry = Assert.Single(remaining.Entries);
                Assert.Equal("Evidence", entry.TermMarkdown);
            });
    }

    [Fact]
    public void HtmlToMarkdown_Can_Expand_Simple_Definition_List_Entries_Via_Transform() {
        var options = new HtmlToMarkdownOptions();
        options.DocumentTransforms.Add(new MarkdownSimpleDefinitionListParagraphTransform());

        var document = """
<dl>
  <dt>Status</dt><dd>healthy</dd>
  <dt>Impact</dt><dd>none</dd>
</dl>
""".LoadFromHtml(options);

        var markdown = NormalizeMarkdown(document.ToMarkdown());
        var html = document.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Equal(NormalizeMarkdown("""
Status: healthy

Impact: none
"""), markdown);
        Assert.DoesNotContain("<dl>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownInlineNormalizationTransform_Is_Idempotent() {
        var transform = new MarkdownInlineNormalizationTransform(new MarkdownInputNormalizationOptions {
            NormalizeTightParentheticalSpacing = true,
            NormalizeTightColonSpacing = true
        });

        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.DocumentTransforms.Add(transform);

        var document = MarkdownReader.Parse("""
Signal **Deleted object remnants**(SID left in ACL path)

Why it matters:missing evidence
""", options);

        var once = NormalizeMarkdown(document.ToMarkdown());
        var twice = NormalizeMarkdown(MarkdownDocumentTransformPipeline.Apply(
            document,
            [transform],
            new MarkdownDocumentTransformContext(MarkdownDocumentTransformSource.MarkdownReader, options)).ToMarkdown());

        Assert.Equal(once, twice);
    }

    [Fact]
    public void MarkdownInlineNormalizationTransform_Updates_Footnote_Text_From_RewrittenBlocks() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownInlineNormalizationTransform(new MarkdownInputNormalizationOptions {
            NormalizeTightColonSpacing = true
        }));

        var document = MarkdownReader.Parse("""
Lead[^1]

[^1]: Why it matters:missing evidence
""", options);

        var footnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(document.Blocks, block => block is FootnoteDefinitionBlock));

        Assert.Equal("Why it matters: missing evidence", footnote.Text);
        Assert.Equal("Why it matters: missing evidence", Assert.Single(footnote.ParagraphBlocks).Inlines.RenderMarkdown());
    }

    [Fact]
    public void MarkdownInlineNormalizationTransform_Updates_Callout_Body_From_RewrittenBlocks() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DocumentTransforms.Add(new MarkdownInlineNormalizationTransform(new MarkdownInputNormalizationOptions {
            NormalizeTightColonSpacing = true
        }));

        var document = MarkdownReader.Parse("""
> [!NOTE] Why it matters
> coverage:missing evidence
""", options);

        var callout = Assert.IsType<CalloutBlock>(Assert.Single(document.Blocks));

        Assert.Equal("coverage: missing evidence", callout.Body);
        Assert.Equal("coverage: missing evidence", Assert.IsType<ParagraphBlock>(Assert.Single(callout.ChildBlocks)).Inlines.RenderMarkdown());
    }

    private static string NormalizeMarkdown(string markdown) {
        return (markdown ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim();
    }

    private sealed class AppendParagraphTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            Assert.Equal(MarkdownDocumentTransformSource.MarkdownReader, context.Source);
            document.Add(new ParagraphBlock(new InlineSequence().Text(text)));
            return document;
        }
    }
}
