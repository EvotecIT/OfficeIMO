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
