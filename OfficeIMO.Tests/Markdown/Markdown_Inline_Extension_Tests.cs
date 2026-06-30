using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Inline_Extension_Tests {
    [Fact]
    public void InlineParserExtensions_Use_Registration_Order_And_First_Success_Wins() {
        var secondClaimStartCalls = 0;
        var options = new MarkdownReaderOptions();
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension("first-claim", TryParseFirstClaimInline));
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension(
            "second-claim",
            (MarkdownInlineParserContext context, out MarkdownInlineParseResult result) => {
                if (IsClaimStart(context)) {
                    secondClaimStartCalls++;
                }

                return TryParseClaimInline(context, "second", out result);
            }));

        var document = MarkdownReader.Parse("Lead {{core}} tail", options);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var custom = Assert.IsType<TestClaimInline>(paragraph.Inlines.Nodes[1]);
        Assert.Equal("first", custom.Value);
        Assert.Equal(0, secondClaimStartCalls);
    }

    [Fact]
    public void InlineParserExtensions_Fall_Back_To_Later_Extensions_When_Parser_Returns_False() {
        var firstCalls = 0;
        var options = new MarkdownReaderOptions();
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension(
            "miss",
            (MarkdownInlineParserContext _, out MarkdownInlineParseResult result) => {
                firstCalls++;
                result = default;
                return false;
            }));
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension("second-claim", TryParseSecondClaimInline));

        var document = MarkdownReader.Parse("Lead {{core}} tail", options);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var custom = Assert.IsType<TestClaimInline>(paragraph.Inlines.Nodes[1]);
        Assert.Equal("second", custom.Value);
        Assert.True(firstCalls > 0);
    }

    [Fact]
    public void InlineParserExtensions_Disabled_Extensions_Are_Skipped_And_Core_Inline_Parsing_Continues() {
        var disabledCalls = 0;
        var options = new MarkdownReaderOptions();
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension(
            "disabled-throwing-extension",
            (MarkdownInlineParserContext _, out MarkdownInlineParseResult result) => {
                disabledCalls++;
                result = default;
                throw new InvalidOperationException("Disabled inline extension should not run.");
            },
            _ => false));

        var document = MarkdownReader.Parse("Lead **bold** tail", options);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        Assert.IsType<BoldSequenceInline>(paragraph.Inlines.Nodes[1]);
        Assert.Equal(0, disabledCalls);
    }

    [Fact]
    public void InlineParserContext_Can_Create_Normalized_SourceSlice_For_Custom_Inline() {
        MarkdownSourceSlice sourceSlice = default;
        bool sourceSliceOk = false;
        var options = new MarkdownReaderOptions();
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension(
            "source-aware-double-brace",
            (MarkdownInlineParserContext context, out MarkdownInlineParseResult result) => {
                result = default;
                if (!IsClaimStart(context)) {
                    return false;
                }

                var closing = context.Text.IndexOf("}}", context.Position + 2, StringComparison.Ordinal);
                if (closing < 0) {
                    return false;
                }

                var consumedLength = closing + 2 - context.Position;
                sourceSliceOk = context.TryCreateSourceSlice(0, consumedLength, out sourceSlice);
                result = new MarkdownInlineParseResult(
                    new TestClaimInline("source-aware"),
                    consumedLength);
                return true;
            }));

        var document = MarkdownReader.Parse("Lead {{core}}\r\n", options);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var custom = Assert.IsType<TestClaimInline>(paragraph.Inlines.Nodes[1]);
        Assert.Equal("source-aware", custom.Value);
        Assert.True(sourceSliceOk);
        Assert.Equal(MarkdownSourceTextKind.Normalized, sourceSlice.TextKind);
        Assert.Equal("{{core}}", sourceSlice.Text);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 13), sourceSlice.SourceSpan);
    }

    [Fact]
    public void InlineTransformExtensions_Run_After_Core_Parsing_In_Registration_Order() {
        var calls = new List<string>();
        var options = new MarkdownReaderOptions();
        options.InlineTransformExtensions.Add(new MarkdownInlineTransformExtension(
            "first",
            (sequence, context) => {
                if (context.IsNestedSequence) {
                    return sequence;
                }

                calls.Add("first:" + sequence.Nodes.Count + ":" + sequence.Nodes[1].GetType().Name);
                sequence.ReplaceItems(new IMarkdownInline[] {
                    new TextRun("one")
                });
                return sequence;
            }));
        options.InlineTransformExtensions.Add(new MarkdownInlineTransformExtension(
            "second",
            (sequence, context) => {
                if (context.IsNestedSequence) {
                    return sequence;
                }

                calls.Add("second:" + sequence.RenderMarkdown());
                sequence.ReplaceItems(new IMarkdownInline[] {
                    new TextRun(sequence.RenderMarkdown() + " two")
                });
                return sequence;
            }));

        var document = MarkdownReader.Parse("Lead **bold**", options);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var text = Assert.IsType<TextRun>(Assert.Single(paragraph.Inlines.Nodes));
        Assert.Equal("one two", text.Text);
        Assert.Equal(new[] { "first:2:BoldSequenceInline", "second:one" }, calls);
    }

    [Fact]
    public void InlineTransformContext_Can_Create_SourceSlices_For_Parsed_Inlines() {
        MarkdownSourceSpan? strongSpan = null;
        MarkdownSourceSlice strongSlice = default;
        bool strongSliceOk = false;
        var options = new MarkdownReaderOptions();
        options.InlineTransformExtensions.Add(new MarkdownInlineTransformExtension(
            "source-aware-transform",
            (sequence, context) => {
                if (context.IsNestedSequence) {
                    return sequence;
                }

                var strong = sequence.Nodes.OfType<BoldSequenceInline>().Single();
                strongSpan = context.GetSourceSpan(strong);
                strongSliceOk = context.TryCreateSourceSlice(strong, out strongSlice);
                return sequence;
            }));

        var document = MarkdownReader.Parse("Lead **Bold**\r\n", options);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        Assert.IsType<BoldSequenceInline>(paragraph.Inlines.Nodes[1]);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 13), strongSpan);
        Assert.True(strongSliceOk);
        Assert.Equal(MarkdownSourceTextKind.Normalized, strongSlice.TextKind);
        Assert.Equal("**Bold**", strongSlice.Text);
    }

    [Fact]
    public void InlineTransformExtensions_Can_Return_Replacement_Sequences_And_Null_Noops() {
        var options = new MarkdownReaderOptions();
        options.InlineTransformExtensions.Add(new MarkdownInlineTransformExtension(
            "replace",
            static (_, _) => new InlineSequence().Text("replacement")));
        options.InlineTransformExtensions.Add(new MarkdownInlineTransformExtension(
            "noop",
            static (_, _) => null));

        var document = MarkdownReader.Parse("original", options);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var text = Assert.IsType<TextRun>(Assert.Single(paragraph.Inlines.Nodes));
        Assert.Equal("replacement", text.Text);
    }

    [Fact]
    public void InlineTransformExtensions_Disabled_Extensions_Are_Skipped() {
        var disabledCalls = 0;
        var options = new MarkdownReaderOptions();
        options.InlineTransformExtensions.Add(new MarkdownInlineTransformExtension(
            "disabled",
            (_, _) => {
                disabledCalls++;
                throw new InvalidOperationException("Disabled inline transform should not run.");
            },
            _ => false));

        var document = MarkdownReader.Parse("Lead **bold** tail", options);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        Assert.IsType<BoldSequenceInline>(paragraph.Inlines.Nodes[1]);
        Assert.Equal(0, disabledCalls);
    }

    [Fact]
    public void InlineTransformExtensions_Visit_Nested_Inline_Containers() {
        var options = new MarkdownReaderOptions();
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension("double-brace", TryParseDoubleBraceInline));
        options.InlineTransformExtensions.Add(new MarkdownInlineTransformExtension(
            "nested-upper",
            static (sequence, context) => {
                if (!context.IsNestedSequence) {
                    return sequence;
                }

                sequence.ReplaceItems(sequence.Nodes.Select(node =>
                    node is TextRun textRun
                        ? new TextRun(textRun.Text.ToUpperInvariant())
                        : node));
                return sequence;
            }));

        var document = MarkdownReader.Parse("Lead {{core}} tail", options);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var custom = Assert.IsType<DoubleBraceInline>(paragraph.Inlines.Nodes[1]);
        var text = Assert.IsType<TextRun>(Assert.Single(custom.Inlines.Nodes));
        Assert.Equal("CORE", text.Text);
    }

    [Fact]
    public void InlineTransformExtensions_Preserve_SourceSpans_For_Reused_Nodes() {
        var options = new MarkdownReaderOptions();
        options.InlineTransformExtensions.Add(new MarkdownInlineTransformExtension(
            "append-spanless-tail",
            static (sequence, context) => {
                if (context.IsNestedSequence) {
                    return sequence;
                }

                var nodes = sequence.Nodes.ToList();
                nodes.Add(new TextRun(" tail"));
                sequence.ReplaceItems(nodes);
                return sequence;
            }));

        var result = MarkdownReader.ParseWithSyntaxTree("Lead **Bold**", options);

        var paragraph = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(new[] {
            MarkdownSyntaxKind.InlineText,
            MarkdownSyntaxKind.InlineStrong,
            MarkdownSyntaxKind.InlineText
        }, paragraph.Children.Select(node => node.Kind).ToArray());
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 13), paragraph.Children[1].SourceSpan);
        Assert.Null(paragraph.Children[2].SourceSpan);
    }

    [Fact]
    public void Syntax_Aware_InlineParserExtension_Projects_Custom_Syntax_To_Native_Snapshots() {
        var options = new MarkdownReaderOptions();
        options.InlineParserExtensions.Add(new MarkdownInlineParserExtension("angle-claim", TryParseAngleClaimInline));

        var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("Lead <<**core**>> tail", options);
        var paragraphSyntax = Assert.Single(result.SyntaxTree.Children);
        var customSyntax = Assert.Single(paragraphSyntax.Children, node => node.CustomKind == "angle-claim-inline");

        Assert.Equal(MarkdownSyntaxKind.Unknown, customSyntax.Kind);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 17), customSyntax.SourceSpan);
        Assert.IsType<AngleClaimInline>(customSyntax.AssociatedObject);
        Assert.Contains(customSyntax.Children, node => node.Kind == MarkdownSyntaxKind.InlineStrong);

        var native = MarkdownNativeDocument.FromParseResult(result);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var custom = Assert.Single(paragraph.InlineRuns, inline => inline.SyntaxKind == MarkdownSyntaxKind.Unknown);

        Assert.Equal(MarkdownNativeInlineKind.Other, custom.Kind);
        Assert.Equal("core", custom.Text);
        Assert.Equal("<<**core**>>", custom.Markdown);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 17), custom.SourceSpan);
        Assert.Single(custom.Children, inline => inline.Kind == MarkdownNativeInlineKind.Strong);

        var snapshot = native.ToSnapshot();
        var snapshotCustom = Assert.Single(snapshot.Blocks[0].Inlines, inline => inline.SyntaxKind == MarkdownSyntaxKind.Unknown);

        Assert.Equal(MarkdownNativeInlineKind.Other, snapshotCustom.Kind);
        Assert.Equal("core", snapshotCustom.Text);
        Assert.Equal("<<**core**>>", snapshotCustom.Markdown);
        Assert.Equal("L1:C6-C17", snapshotCustom.SourceSpan!.Display);
        Assert.Single(snapshotCustom.Children, inline => inline.Kind == MarkdownNativeInlineKind.Strong);
    }

    [Fact]
    public void InlineMarkdownRenderExtensions_Use_Last_Registration_And_Null_Fallback() {
        var options = new MarkdownWriteOptions {
            OutputLineEnding = "\n"
        };
        options.InlineRenderExtensions.Add(new MarkdownInlineMarkdownRenderExtension(
            "first",
            typeof(TestClaimInline),
            static (inline, _) => "[first:" + ((TestClaimInline)inline).Value + "]"));
        options.InlineRenderExtensions.Add(new MarkdownInlineMarkdownRenderExtension(
            "last",
            typeof(TestClaimInline),
            static (inline, _) => "[last:" + ((TestClaimInline)inline).Value + "]"));

        var document = MarkdownDoc.Create()
            .Add(new ParagraphBlock(new InlineSequence()
                .Text("Lead")
                .AddRaw(new TestClaimInline("core"))
                .Text("tail")));

        Assert.Equal("Lead [last:core] tail\n", document.ToMarkdown(options));
    }

    [Fact]
    public void InlineMarkdownRenderExtensions_Contextual_Renderer_Can_Read_Document_Context_And_SourceSpan() {
        var readerOptions = new MarkdownReaderOptions();
        readerOptions.InlineParserExtensions.Add(new MarkdownInlineParserExtension("double-brace", TryParseDoubleBraceInline));
        var document = MarkdownReader.ParseWithSyntaxTree("""
## Intro

Lead {{core}} tail
""", readerOptions).Document;

        var options = new MarkdownWriteOptions {
            OutputLineEnding = "\n"
        };
        options.InlineRenderExtensions.Add(MarkdownInlineMarkdownRenderExtension.CreateContextual(
            "double-brace-contextual-markdown",
            typeof(DoubleBraceInline),
            static (inline, context) => {
                if (inline is not DoubleBraceInline custom) {
                    return null;
                }

                var paragraph = custom.Ancestors().OfType<ParagraphBlock>().First();
                var syntax = context.FindSyntaxNode(custom);
                var hasSourceSlice = context.TryCreateSourceSlice(custom, out var sourceSlice);
                return "[block:"
                    + context.GetBlockIndex(paragraph)
                    + ";kind:"
                    + syntax?.Kind
                    + ";source:"
                    + (hasSourceSlice ? sourceSlice.Text : string.Empty)
                    + ";text:"
                    + InlinePlainText.Extract(custom.Inlines)
                    + "]";
            }));

        var rendered = document.ToMarkdown(options);

        Assert.Equal("## Intro\n\nLead [block:1;kind:Unknown;source:{{core}};text:core] tail\n", rendered);
    }

    [Fact]
    public void InlineMarkdownRenderExtensions_Fall_Back_When_Renderer_Returns_Null() {
        var options = new MarkdownWriteOptions {
            OutputLineEnding = "\n"
        };
        options.InlineRenderExtensions.Add(new MarkdownInlineMarkdownRenderExtension(
            "fallback",
            typeof(TestClaimInline),
            static (_, _) => null));

        var document = MarkdownDoc.Create()
            .Add(new ParagraphBlock(new InlineSequence()
                .Text("Lead")
                .AddRaw(new TestClaimInline("core"))
                .Text("tail")));

        Assert.Equal("Lead core tail\n", document.ToMarkdown(options));
    }

    private static bool TryParseFirstClaimInline(MarkdownInlineParserContext context, out MarkdownInlineParseResult result) =>
        TryParseClaimInline(context, "first", out result);

    private static bool TryParseSecondClaimInline(MarkdownInlineParserContext context, out MarkdownInlineParseResult result) =>
        TryParseClaimInline(context, "second", out result);

    private static bool TryParseClaimInline(
        MarkdownInlineParserContext context,
        string value,
        out MarkdownInlineParseResult result) {
        result = default;
        if (!IsClaimStart(context)) {
            return false;
        }

        var closing = context.Text.IndexOf("}}", context.Position + 2, StringComparison.Ordinal);
        if (closing < 0) {
            return false;
        }

        result = new MarkdownInlineParseResult(
            new TestClaimInline(value),
            closing + 2 - context.Position);
        return true;
    }

    private static bool IsClaimStart(MarkdownInlineParserContext context) =>
        context.CurrentChar == '{'
        && context.Position + 1 < context.Text.Length
        && context.Text[context.Position + 1] == '{';

    private static bool TryParseDoubleBraceInline(MarkdownInlineParserContext context, out MarkdownInlineParseResult result) {
        result = default;
        if (!IsClaimStart(context)) {
            return false;
        }

        var closing = context.Text.IndexOf("}}", context.Position + 2, StringComparison.Ordinal);
        if (closing < 0) {
            return false;
        }

        var innerLength = closing - (context.Position + 2);
        result = new MarkdownInlineParseResult(
            new DoubleBraceInline(context.ParseNestedInlines(2, innerLength)),
            closing + 2 - context.Position);
        return true;
    }

    private static bool TryParseAngleClaimInline(MarkdownInlineParserContext context, out MarkdownInlineParseResult result) {
        result = default;
        if (context.CurrentChar != '<'
            || context.Position + 1 >= context.Text.Length
            || context.Text[context.Position + 1] != '<') {
            return false;
        }

        var closing = context.Text.IndexOf(">>", context.Position + 2, StringComparison.Ordinal);
        if (closing < 0) {
            return false;
        }

        var innerLength = closing - (context.Position + 2);
        result = new MarkdownInlineParseResult(
            new AngleClaimInline(context.ParseNestedInlines(2, innerLength)),
            closing + 2 - context.Position);
        return true;
    }

    private sealed class TestClaimInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
        public TestClaimInline(string value) {
            Value = value;
        }

        public string Value { get; }

        string IRenderableMarkdownInline.RenderMarkdown() => Value;

        string IRenderableMarkdownInline.RenderHtml() => System.Net.WebUtility.HtmlEncode(Value);

        void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Value);
    }

    private sealed class DoubleBraceInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IInlineContainerMarkdownInline {
        public DoubleBraceInline(InlineSequence inlines) {
            Inlines = inlines;
        }

        public InlineSequence Inlines { get; }

        InlineSequence? IInlineContainerMarkdownInline.NestedInlines => Inlines;

        string IRenderableMarkdownInline.RenderMarkdown() => "{{" + Inlines.RenderMarkdown() + "}}";

        string IRenderableMarkdownInline.RenderHtml() => "<span>" + Inlines.RenderHtml() + "</span>";

        void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => InlinePlainText.AppendPlainText(sb, Inlines);
    }

    private sealed class AngleClaimInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IInlineContainerMarkdownInline, ISyntaxMarkdownInline {
        public AngleClaimInline(InlineSequence inlines) {
            Inlines = inlines;
        }

        public InlineSequence Inlines { get; }

        InlineSequence? IInlineContainerMarkdownInline.NestedInlines => Inlines;

        string IRenderableMarkdownInline.RenderMarkdown() => "<<" + Inlines.RenderMarkdown() + ">>";

        string IRenderableMarkdownInline.RenderHtml() => "<span data-angle-claim=\"true\">" + Inlines.RenderHtml() + "</span>";

        void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => InlinePlainText.AppendPlainText(sb, Inlines);

        MarkdownSyntaxNode ISyntaxMarkdownInline.BuildSyntaxNode(MarkdownInlineSyntaxBuilderContext context, MarkdownSourceSpan? span) {
            var children = context.BuildChildren(Inlines);
            return new MarkdownSyntaxNode(
                MarkdownSyntaxKind.Unknown,
                span ?? context.GetAggregateSpan(children),
                ((IRenderableMarkdownInline)this).RenderMarkdown(),
                children,
                this,
                "angle-claim-inline");
        }
    }
}
