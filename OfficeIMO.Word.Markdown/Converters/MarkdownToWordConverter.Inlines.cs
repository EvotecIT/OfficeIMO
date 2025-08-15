using Markdig.Extensions.Footnotes;
using Markdig.Renderers.Html;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using SixLabors.ImageSharp;

namespace OfficeIMO.Word.Markdown.Converters {
    internal partial class MarkdownToWordConverter {
        private static void ProcessInline(Inline? inline, WordParagraph paragraph, MarkdownToWordOptions options, WordDocument document) {
            if (inline == null) {
                return;
            }

            void Handle(Inline? node, bool bold = false, bool italic = false, bool strike = false) {
                for (var current = node; current != null; current = current.NextSibling) {
                    switch (current) {
                        case LiteralInline literal:
                            var run = paragraph.AddText(literal.Content.ToString());
                            if (bold) run.SetBold();
                            if (italic) run.SetItalic();
                            if (strike) run.SetStrike();
                            if (!string.IsNullOrEmpty(options.FontFamily)) {
                                run.SetFontFamily(options.FontFamily);
                            }
                            break;
                        case EmphasisInline emphasis:
                            bool eBold = bold;
                            bool eItalic = italic;
                            bool eStrike = strike;
                            if (emphasis.DelimiterChar == '~') {
                                eStrike = true;
                            } else {
                                if (emphasis.DelimiterCount == 1) {
                                    eItalic = true;
                                } else if (emphasis.DelimiterCount == 2) {
                                    eBold = true;
                                } else if (emphasis.DelimiterCount >= 3) {
                                    eBold = true;
                                    eItalic = true;
                                }
                            }
                            Handle(emphasis.FirstChild, eBold, eItalic, eStrike);
                            break;
                        case LinkInline link:
                            if (link.IsImage) {
                                AddImage(document, paragraph, link);
                            } else {
                                string label = GetPlainText(link.FirstChild);
                                var hyperlink = paragraph.AddHyperLink(label, new Uri(link.Url, UriKind.RelativeOrAbsolute));
                                if (!string.IsNullOrEmpty(options.FontFamily)) {
                                    hyperlink.SetFontFamily(options.FontFamily);
                                }
                            }
                            break;
                        case FootnoteLink footnoteLink:
                            string text = BuildFootnoteText(footnoteLink.Footnote);
                            paragraph.AddFootNote(text);
                            break;
                        case LineBreakInline:
                            paragraph.AddBreak();
                            break;
                        case ContainerInline container:
                            Handle(container.FirstChild, bold, italic, strike);
                            break;
                        default:
                            if (current is LeafInline leaf) {
                                var other = paragraph.AddText(leaf.ToString());
                                if (bold) other.SetBold();
                                if (italic) other.SetItalic();
                                if (strike) other.SetStrike();
                                if (!string.IsNullOrEmpty(options.FontFamily)) {
                                    other.SetFontFamily(options.FontFamily);
                                }
                            }
                            break;
                    }
                }
            }

            var start = inline is ContainerInline container ? container.FirstChild : inline;
            Handle(start);
        }

        private static void AddImage(WordDocument document, WordParagraph paragraph, LinkInline link) {
            string url = link.Url?.Trim() ?? string.Empty;
            string altText = GetPlainText(link.FirstChild);
            string? title = link.Title?.Trim();
            double? width = null;
            double? height = null;
            byte[]? imageData = null;
            string? remoteFileName = null;

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

            if (width == null && height == null) {
                try {
                    if (url.StartsWith("http", StringComparison.OrdinalIgnoreCase)) {
                        using HttpClient client = new();
                        imageData = client.GetByteArrayAsync(url).GetAwaiter().GetResult();
                        using var image = Image.Load(imageData, out var format);
                        width = image.Width;
                        height = image.Height;
                        string extension = format.FileExtensions.FirstOrDefault() ?? "png";
                        try {
                            remoteFileName = Path.GetFileName(new Uri(url).LocalPath);
                        } catch {
                            remoteFileName = null;
                        }
                        if (string.IsNullOrEmpty(remoteFileName)) {
                            remoteFileName = "image." + extension;
                        } else if (string.IsNullOrEmpty(Path.GetExtension(remoteFileName))) {
                            remoteFileName += "." + extension;
                        }
                    } else if (File.Exists(url)) {
                        using var image = Image.Load(url, out _);
                        width = image.Width;
                        height = image.Height;
                    }
                } catch {
                    // ignore errors when determining natural image size
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
                if (imageData != null) {
                    using var ms = new MemoryStream(imageData);
                    paragraph.AddImage(ms, remoteFileName ?? "image", width, height, description: altText);
                } else {
                    var img = document.AddImageFromUrl(url, width, height);
                    if (!string.IsNullOrEmpty(altText)) {
                        img.Description = altText;
                    }
                }
            } else {
                paragraph.AddImage(url, width, height, description: altText);
            }
        }

        private static string BuildFootnoteText(Footnote footnote) {
            var sb = new StringBuilder();
            bool first = true;
            foreach (var block in footnote) {
                if (block is ParagraphBlock pb) {
                    if (!first) {
                        sb.AppendLine();
                    }
                    sb.Append(GetPlainText(pb.Inline));
                    first = false;
                }
            }
            return sb.ToString();
        }

        private static string GetPlainText(Inline? inline) {
            if (inline == null) {
                return string.Empty;
            }

            var sb = new StringBuilder();
            void Append(Inline? node) {
                for (var current = node; current != null; current = current.NextSibling) {
                    switch (current) {
                        case LiteralInline literal:
                            sb.Append(literal.Content.ToString());
                            break;
                        case EmphasisInline emphasis:
                            Append(emphasis.FirstChild);
                            break;
                        case LineBreakInline:
                            sb.Append('\n');
                            break;
                        case ContainerInline container:
                            Append(container.FirstChild);
                            break;
                    }
                }
            }

            Append(inline);
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