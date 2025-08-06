using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Markdown.Converters {
    internal partial class MarkdownToWordConverter {
        private static void ProcessInline(Inline? inline, WordParagraph paragraph, MarkdownToWordOptions options, WordDocument document) {
            if (inline == null) {
                return;
            }

            var buffer = new StringBuilder();

            void Flush() {
                if (buffer.Length > 0) {
                    InlineRunHelper.AddInlineRuns(paragraph, buffer.ToString(), options.FontFamily);
                    buffer.Clear();
                }
            }

            for (var current = inline; current != null; current = current.NextSibling) {
                if (current is LinkInline link) {
                    Flush();
                    if (link.IsImage) {
                        AddImage(document, paragraph, link);
                    } else {
                        string label = BuildMarkdown(link.FirstChild);
                        var hyperlink = paragraph.AddHyperLink(label, new Uri(link.Url, UriKind.RelativeOrAbsolute));
                        if (!string.IsNullOrEmpty(options.FontFamily)) {
                            hyperlink.SetFontFamily(options.FontFamily);
                        }
                    }
                } else {
                    buffer.Append(BuildMarkdown(current));
                }
            }
            Flush();
        }

        private static void AddImage(WordDocument document, WordParagraph paragraph, LinkInline link) {
            if (File.Exists(link.Url)) {
                paragraph.AddImage(link.Url);
            } else {
                document.AddImageFromUrl(link.Url, 50, 50);
            }
        }

        private static string BuildMarkdown(Inline? inline) {
            if (inline == null) {
                return string.Empty;
            }

            var sb = new StringBuilder();
            for (var current = inline; current != null; current = current.NextSibling) {
                switch (current) {
                    case LiteralInline literal:
                        sb.Append(literal.Content.ToString());
                        break;
                    case EmphasisInline emphasis:
                        string marker = new('*', emphasis.DelimiterCount);
                        sb.Append(marker);
                        sb.Append(BuildMarkdown(emphasis.FirstChild));
                        sb.Append(marker);
                        break;
                    case LineBreakInline:
                        sb.Append('\n');
                        break;
                    case ContainerInline container:
                        sb.Append(BuildMarkdown(container.FirstChild));
                        break;
                }
            }

            return sb.ToString();
        }

        private static string GetCodeBlockText(CodeBlock codeBlock) {
            var sb = new StringBuilder();
            foreach (var line in codeBlock.Lines.Lines) {
                sb.AppendLine(line.Slice.ToString());
            }
            return sb.ToString().TrimEnd();
        }
    }
}