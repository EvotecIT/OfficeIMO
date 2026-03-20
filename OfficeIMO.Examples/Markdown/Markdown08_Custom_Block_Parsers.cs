using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static class Markdown08_Custom_Block_Parsers {
        public static void Example_Custom_Block_Parsers(string folderPath, bool open = false) {
            Console.WriteLine("[*] Markdown custom block parser extensions");

            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);

            string markdown = """
# Panel extension demo

:::panel Release Notes
This custom block uses the delegate-based block parser API.

- Nested markdown is parsed with preserved source spans
- Child blocks participate in traversal and syntax generation
:::
""";

            var options = CreateReaderOptions();
            var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
            var html = result.Document.ToHtmlFragment(new HtmlOptions {
                Kind = HtmlKind.Fragment,
                Title = "Custom block parser sample",
                Style = HtmlStyle.GithubAuto
            });

            string markdownPath = Path.Combine(mdFolder, "CustomBlockParsers.Source.md");
            string htmlPath = Path.Combine(mdFolder, "CustomBlockParsers.html");
            string syntaxPath = Path.Combine(mdFolder, "CustomBlockParsers.SyntaxTree.txt");

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
            var options = MarkdownReaderOptions.CreatePortableProfile();
            options.BlockParserExtensions.Add(new MarkdownBlockParserExtension(
                "panel-block",
                MarkdownBlockParserPlacement.BeforeParagraphs,
                TryParsePanelBlock));
            return options;
        }

        private static string DescribeParseResult(MarkdownParseResult result) {
            var sb = new StringBuilder();
            sb.AppendLine("Document blocks:");
            foreach (var block in result.Document.Blocks) {
                sb.AppendLine("- " + block.GetType().Name);
            }

            sb.AppendLine();
            sb.AppendLine("Nested panel descendants:");
            foreach (var paragraph in result.Document.DescendantsOfType<ParagraphBlock>()) {
                sb.AppendLine($"- Paragraph @ {FormatSpan(paragraph.SourceSpan)}");
            }

            sb.AppendLine();
            sb.AppendLine("Syntax tree:");
            AppendSyntaxNode(sb, result.SyntaxTree, 0);
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

        private static bool TryParsePanelBlock(MarkdownBlockParserContext context, out MarkdownBlockParseResult result) {
            result = default;
            const string prefix = ":::panel";
            var trimmed = context.CurrentLine.Trim();
            if (!trimmed.StartsWith(prefix, StringComparison.Ordinal)) {
                return false;
            }

            var title = trimmed.Length == prefix.Length ? string.Empty : trimmed.Substring(prefix.Length).Trim();
            var closingOffset = -1;
            for (var offset = 1; context.TryGetLine(offset, out var line); offset++) {
                if (string.Equals(line.Trim(), ":::", StringComparison.Ordinal)) {
                    closingOffset = offset;
                    break;
                }
            }

            if (closingOffset < 0) {
                return false;
            }

            var nestedBlocks = closingOffset > 1
                ? context.ParseNestedBlocks(1, closingOffset - 1)
                : Array.Empty<IMarkdownBlock>();
            result = new MarkdownBlockParseResult(new PanelBlock(title, nestedBlocks), closingOffset + 1);
            return true;
        }

        private sealed class PanelBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlockWithContext, IContextualHtmlMarkdownBlock, IChildMarkdownBlockContainer {
            public PanelBlock(string title, IReadOnlyList<IMarkdownBlock> childBlocks) {
                Title = title ?? string.Empty;
                ChildBlocks = childBlocks ?? Array.Empty<IMarkdownBlock>();
            }

            public string Title { get; }
            public IReadOnlyList<IMarkdownBlock> ChildBlocks { get; }

            string IMarkdownBlock.RenderMarkdown() {
                var body = string.Join("\n\n", ChildBlocks.Select(block => block.RenderMarkdown().TrimEnd()));
                return string.IsNullOrWhiteSpace(body)
                    ? $":::panel {Title}\n:::"
                    : $":::panel {Title}\n{body}\n:::";
            }

            string IMarkdownBlock.RenderHtml() {
                var body = string.Concat(ChildBlocks.Select(block => block.RenderHtml()));
                return $"<section data-panel=\"true\"><h2>{System.Net.WebUtility.HtmlEncode(Title)}</h2>{body}</section>";
            }

            string IContextualHtmlMarkdownBlock.RenderHtml(MarkdownBodyRenderContext context) {
                var body = string.Concat(ChildBlocks.Select(block => block.RenderHtml()));
                var blockIndex = context.GetBlockIndex(this);
                var tocOptions = new TocOptions {
                    Scope = TocScope.PreviousHeading,
                    IncludeTitle = true,
                    MinLevel = 2,
                    MaxLevel = 6
                };
                var titleAnchor = context.GetPrecedingHeadingAnchor(blockIndex, tocOptions);
                var entries = context.BuildTocEntries(blockIndex, tocOptions, titleAnchor);
                var nav = string.Concat(entries.Select(entry =>
                    $"<a href=\"#{System.Net.WebUtility.HtmlEncode(entry.Anchor)}\">{System.Net.WebUtility.HtmlEncode(entry.Text)}</a>"));
                return $"<section data-panel=\"true\" data-title=\"{System.Net.WebUtility.HtmlEncode(context.Options.Title)}\" data-block-index=\"{blockIndex}\" data-parent-anchor=\"{System.Net.WebUtility.HtmlEncode(titleAnchor)}\"><h2>{System.Net.WebUtility.HtmlEncode(Title)}</h2>{nav}{body}</section>";
            }

            public MarkdownSyntaxNode BuildSyntaxNode(MarkdownBlockSyntaxBuilderContext context, MarkdownSourceSpan? span) {
                var titleNode = context.BuildInlineContainerNode(
                    MarkdownSyntaxKind.Paragraph,
                    new InlineSequence().Text(Title),
                    literal: Title);
                var childNodes = context.BuildChildSyntaxNodes(ChildBlocks);
                var children = new List<MarkdownSyntaxNode>(childNodes.Count + 1) { titleNode };
                for (int i = 0; i < childNodes.Count; i++) {
                    children.Add(childNodes[i]);
                }

                return new MarkdownSyntaxNode(
                    MarkdownSyntaxKind.Unknown,
                    span ?? context.GetAggregateSpan(children),
                    literal: context.NormalizeLiteralLineEndings(((IMarkdownBlock)this).RenderMarkdown()),
                    children: children,
                    associatedObject: this,
                    customKind: "panel-block");
            }
        }
    }
}
