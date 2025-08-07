using Markdig.Renderers.Html;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

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

            var start = inline is ContainerInline container ? container.FirstChild : inline;
            for (var current = start; current != null; current = current.NextSibling) {
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
                } else if (current is EmphasisInline emphasis && emphasis.DelimiterChar == '~') {
                    Flush();
                    string text = BuildMarkdown(emphasis.FirstChild);
                    var run = paragraph.AddFormattedText(text);
                    run.SetStrike();
                    if (!string.IsNullOrEmpty(options.FontFamily)) {
                        run.SetFontFamily(options.FontFamily);
                    }
                } else {
                    buffer.Append(BuildMarkdown(current));
                }
            }
            Flush();
        }

        private static void AddImage(WordDocument document, WordParagraph paragraph, LinkInline link) {
            string url = link.Url?.Trim() ?? string.Empty;
            string? title = link.Title?.Trim();
            double? width = null;
            double? height = null;

            // Check for size hints in URL (e.g. "path =100x200")
            var matchUrl = Regex.Match(url, @"\s*=([0-9]+)(?:x([0-9]+))?\s*$");
            if (matchUrl.Success) {
                width = double.Parse(matchUrl.Groups[1].Value);
                if (matchUrl.Groups[2].Success) {
                    height = double.Parse(matchUrl.Groups[2].Value);
                }
                url = url.Substring(0, matchUrl.Index).Trim();
            }

            // Size hints may also appear after the title
            if (!string.IsNullOrEmpty(title)) {
                var matchTitle = Regex.Match(title, @"\s*=([0-9]+)(?:x([0-9]+))?\s*$");
                if (matchTitle.Success) {
                    width ??= double.Parse(matchTitle.Groups[1].Value);
                    if (matchTitle.Groups[2].Success) {
                        height ??= double.Parse(matchTitle.Groups[2].Value);
                    }
                    title = title.Substring(0, matchTitle.Index).Trim();
                }
            }

            // Try to read dimensions from generic attributes
            var attrs = link.GetAttributes();
            if (attrs.Properties != null) {
                if (width == null) {
                    var wProp = attrs.Properties.Find(p => string.Equals(p.Key, "width", StringComparison.OrdinalIgnoreCase));
                    if (wProp.Key != null && double.TryParse(wProp.Value, out var w)) {
                        width = w;
                    }
                }
                if (height == null) {
                    var hProp = attrs.Properties.Find(p => string.Equals(p.Key, "height", StringComparison.OrdinalIgnoreCase));
                    if (hProp.Key != null && double.TryParse(hProp.Value, out var h)) {
                        height = h;
                    }
                }
            }

            if (width == null && height != null) {
                width = height;
            } else if (height == null && width != null) {
                height = width;
            }

            width ??= 50;
            height ??= 50;

            if (url.StartsWith("http", StringComparison.OrdinalIgnoreCase)) {
                var img = document.AddImageFromUrl(url, width, height);
                if (!string.IsNullOrEmpty(title)) {
                    img.Description = title;
                }
            } else {
                paragraph.AddImage(url, width, height, description: title ?? string.Empty);
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
                        char delimiter = emphasis.DelimiterChar;
                        string marker = new(delimiter, emphasis.DelimiterCount);
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