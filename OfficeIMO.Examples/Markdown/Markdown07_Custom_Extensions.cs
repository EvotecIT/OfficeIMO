using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown07_Custom_Extensions {
        public static void Example_Custom_Extensions(string folderPath, bool open = false) {
            Console.WriteLine("[*] Markdown custom extensions");

            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);

            string markdown = """
# Extension demo

Lead {{**Bold** core}} tail

```vendor-chart
{"type":"bar","series":[3,5,2]}
```
""";

            var options = CreateReaderOptions();
            var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
            var html = result.Document.ToHtmlFragment(new HtmlOptions {
                Kind = HtmlKind.Fragment,
                Title = "Custom extension sample",
                Style = HtmlStyle.GithubAuto
            });

            string markdownPath = Path.Combine(mdFolder, "CustomExtensions.Source.md");
            string htmlPath = Path.Combine(mdFolder, "CustomExtensions.html");
            string syntaxPath = Path.Combine(mdFolder, "CustomExtensions.SyntaxTree.txt");

            File.WriteAllText(markdownPath, markdown, Encoding.UTF8);
            File.WriteAllText(htmlPath, html, Encoding.UTF8);
            File.WriteAllText(syntaxPath, DescribeParseResult(result), Encoding.UTF8);

            Console.WriteLine($"✓ Markdown saved: {markdownPath}");
            Console.WriteLine($"✓ HTML saved: {htmlPath}");
            Console.WriteLine($"✓ Syntax tree saved: {syntaxPath}");

            if (open) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(htmlPath) { UseShellExecute = true });
            }
        }

        private static MarkdownReaderOptions CreateReaderOptions() {
            var options = new MarkdownReaderOptions();
            options.InlineParserExtensions.Add(new MarkdownInlineParserExtension("double-brace", TryParseDoubleBraceInline));
            options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
                "Vendor charts",
                new[] { "vendor-chart" },
                context => new VendorChartBlock(context.Language, context.Content)));
            return options;
        }

        private static string DescribeParseResult(MarkdownParseResult result) {
            var sb = new StringBuilder();
            sb.AppendLine("Document blocks:");
            foreach (var block in result.Document.Blocks) {
                sb.AppendLine("- " + block.GetType().Name);
            }

            sb.AppendLine();
            sb.AppendLine("Syntax tree:");
            AppendSyntaxNode(sb, result.SyntaxTree, 0);

            var customInline = result.SyntaxTree.Descendants()
                .FirstOrDefault(node => string.Equals(node.CustomKind, "double-brace", StringComparison.Ordinal));
            var customBlock = result.SyntaxTree.Descendants()
                .FirstOrDefault(node => string.Equals(node.CustomKind, "vendor-chart", StringComparison.Ordinal));

            if (customInline != null) {
                sb.AppendLine();
                sb.AppendLine("Custom inline:");
                sb.AppendLine($"- Kind: {customInline.Kind}");
                sb.AppendLine($"- CustomKind: {customInline.CustomKind}");
                sb.AppendLine($"- Span: {FormatSpan(customInline.SourceSpan)}");
                sb.AppendLine($"- Literal: {customInline.Literal}");
            }

            if (customBlock != null) {
                sb.AppendLine();
                sb.AppendLine("Custom fenced block:");
                sb.AppendLine($"- Kind: {customBlock.Kind}");
                sb.AppendLine($"- CustomKind: {customBlock.CustomKind}");
                sb.AppendLine($"- Span: {FormatSpan(customBlock.SourceSpan)}");
                sb.AppendLine($"- Literal: {customBlock.Literal}");
            }

            return sb.ToString();
        }

        private static void AppendSyntaxNode(StringBuilder sb, MarkdownSyntaxNode node, int depth) {
            sb.Append(' ', depth * 2);
            sb.Append("- ");
            sb.Append(node.Kind);
            if (!string.IsNullOrWhiteSpace(node.CustomKind)) {
                sb.Append(" [");
                sb.Append(node.CustomKind);
                sb.Append(']');
            }
            if (node.SourceSpan.HasValue) {
                sb.Append(" @ ");
                sb.Append(FormatSpan(node.SourceSpan));
            }
            if (!string.IsNullOrWhiteSpace(node.Literal)) {
                sb.Append(" => ");
                sb.Append(node.Literal.Replace("\r", "\\r").Replace("\n", "\\n"));
            }
            sb.AppendLine();

            foreach (var child in node.Children) {
                AppendSyntaxNode(sb, child, depth + 1);
            }
        }

        private static string FormatSpan(MarkdownSourceSpan? span) {
            return span.HasValue
                ? $"{span.Value.StartLine}:{span.Value.StartColumn}-{span.Value.EndLine}:{span.Value.EndColumn}"
                : "(none)";
        }

        private sealed class DoubleBraceInline(InlineSequence inlines) : MarkdownInline, IRenderableMarkdownInline, IContextualHtmlMarkdownInline, IPlainTextMarkdownInline, IInlineContainerMarkdownInline, ISyntaxMarkdownInline {
            public InlineSequence Inlines { get; } = inlines ?? new InlineSequence();

            public string RenderMarkdown() => "{{" + ((IRenderableMarkdownInline)Inlines).RenderMarkdown() + "}}";

            public string RenderHtml() => "<span data-inline=\"double-brace\">" + ((IRenderableMarkdownInline)Inlines).RenderHtml() + "</span>";

            string IContextualHtmlMarkdownInline.RenderHtml(HtmlOptions options) =>
                "<span data-inline=\"double-brace\" data-title=\""
                + System.Net.WebUtility.HtmlEncode(options.Title)
                + "\">"
                + ((IRenderableMarkdownInline)Inlines).RenderHtml()
                + "</span>";

            public void AppendPlainText(StringBuilder sb) => ((IPlainTextMarkdownInline)Inlines).AppendPlainText(sb);

            public MarkdownSyntaxNode BuildSyntaxNode(MarkdownInlineSyntaxBuilderContext context, MarkdownSourceSpan? span) {
                var children = context.BuildChildren(Inlines);
                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.Unknown,
                    span ?? context.GetAggregateSpan(children),
                    literal: RenderMarkdown(),
                    children: children,
                    associatedObject: this,
                    customKind: "double-brace");
            }

            InlineSequence? IInlineContainerMarkdownInline.NestedInlines => Inlines;
        }

        private sealed class VendorChartBlock(string language, string payload) : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlockWithContext, IContextualHtmlMarkdownBlock {
            public string Language { get; } = language ?? string.Empty;
            public string Payload { get; } = payload ?? string.Empty;

            string IMarkdownBlock.RenderMarkdown() => $"```{Language}\n{Payload}\n```";

            string IMarkdownBlock.RenderHtml() =>
                $"<pre><code class=\"language-{System.Net.WebUtility.HtmlEncode(Language)}\">{System.Net.WebUtility.HtmlEncode(Payload)}</code></pre>";

            string IContextualHtmlMarkdownBlock.RenderHtml(MarkdownBodyRenderContext context) =>
                $"<div data-vendor-chart=\"true\" data-title=\"{System.Net.WebUtility.HtmlEncode(context.Options.Title)}\" data-block-count=\"{context.Blocks.Count}\">{System.Net.WebUtility.HtmlEncode(Payload)}</div>";

            public MarkdownSyntaxNode BuildSyntaxNode(MarkdownBlockSyntaxBuilderContext context, MarkdownSourceSpan? span) {
                var payloadNode = context.BuildInlineContainerNode(
                    MarkdownSyntaxKind.Paragraph,
                    new InlineSequence().Text(Payload),
                    literal: Payload);

                var children = new[] {
                    new MarkdownSyntaxNode(MarkdownSyntaxKind.CodeFenceInfo, literal: Language),
                    payloadNode
                };

                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.Unknown,
                    span ?? context.GetAggregateSpan(children),
                    literal: context.NormalizeLiteralLineEndings(((IMarkdownBlock)this).RenderMarkdown()),
                    children: children,
                    associatedObject: this,
                    customKind: "vendor-chart");
            }
        }

        private static bool TryParseDoubleBraceInline(MarkdownInlineParserContext context, out MarkdownInlineParseResult result) {
            result = default;
            if (context.CurrentChar != '{'
                || context.Position + 1 >= context.Text.Length
                || context.Text[context.Position + 1] != '{') {
                return false;
            }

            var closing = context.Text.IndexOf("}}", context.Position + 2, StringComparison.Ordinal);
            if (closing < 0) {
                return false;
            }

            var innerLength = closing - (context.Position + 2);
            var nested = context.ParseNestedInlines(2, innerLength);
            result = new MarkdownInlineParseResult(new DoubleBraceInline(nested), closing + 2 - context.Position);
            return true;
        }
    }
}
