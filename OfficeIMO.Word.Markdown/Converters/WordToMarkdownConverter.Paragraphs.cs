using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word.Markdown {
    internal partial class WordToMarkdownConverter {
        private string ConvertParagraph(WordParagraph paragraph, WordToMarkdownOptions options) {
            const string codeLangPrefix = "CodeLang_";
            string? styleId = paragraph.StyleId;
            string? codeFont = options.FontFamily ?? FontResolver.Resolve("monospace");
            if (styleId is { Length: > 0 } sid && sid.StartsWith(codeLangPrefix, StringComparison.Ordinal) && !string.IsNullOrEmpty(codeFont)) {
                var codeFontValue = codeFont!;
                var runs = paragraph.GetRuns().ToList();
                if (runs.Count > 0 && runs.All(r => string.Equals(r.FontFamily ?? string.Empty, codeFontValue, StringComparison.OrdinalIgnoreCase))) {
                    string language = sid.Substring(codeLangPrefix.Length);
                    string code = string.Concat(runs.Select(r => r.Text));
                    return $"```{language}\n{code}\n```";
                }
            }

            var sb = new StringBuilder();

            if (paragraph.IndentationBefore.HasValue && paragraph.IndentationBefore.Value > 0) {
                int depth = (int)Math.Round(paragraph.IndentationBefore.Value / 720d);
                if (depth > 0) {
                    sb.Append(string.Join(" ", Enumerable.Repeat(">", depth))).Append(' ');
                }
            }

            int? headingLevel = paragraph.Style.HasValue
                ? HeadingStyleMapper.GetLevelForHeadingStyle(paragraph.Style.Value)
                : (int?)null;
            if (headingLevel.HasValue && headingLevel.Value > 0) {
                sb.Append(new string('#', headingLevel.Value)).Append(' ');
            }

            var listInfo = DocumentTraversal.GetListInfo(paragraph);
            if (listInfo != null) {
                int level = listInfo.Value.Level;
                sb.Append(new string(' ', level * 2));
                sb.Append(listInfo.Value.Ordered ? "1. " : "- ");
                // Task list (checkbox) mapping
                if (paragraph.IsCheckBox) {
                    bool done = paragraph.CheckBox?.IsChecked == true;
                    sb.Append(done ? "[x] " : "[ ] ");
                }
            }

            sb.Append(RenderRuns(paragraph, options));

            return sb.ToString();
        }

        private string RenderRuns(WordParagraph paragraph, WordToMarkdownOptions options) {
            var sb = new StringBuilder();
            string? codeFont = options.FontFamily ?? FontResolver.Resolve("monospace");
            foreach (var run in paragraph.GetRuns()) {
                if (run.IsFootNote && run.FootNote != null && run.FootNote.ReferenceId.HasValue) {
                    long id = run.FootNote.ReferenceId.Value;
                    sb.Append($"[^{id}]");
                    continue;
                }

                if (run.IsImage && run.Image != null) {
                    sb.Append(RenderImage(run.Image, options));
                    continue;
                }

                string? text = run.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                if (run.Bold && run.Italic) {
                    text = $"***{text}***";
                } else if (run.Bold) {
                    text = $"**{text}**";
                } else if (run.Italic) {
                    text = $"*{text}*";
                }

                if (options.EnableUnderline && run.Underline.HasValue && run.Underline.Value != UnderlineValues.None) {
                    text = $"<u>{text}</u>";
                }

                if (run.Strike) {
                    text = $"~~{text}~~";
                }

                if (options.EnableHighlight && run.Highlight.HasValue && run.Highlight.Value != HighlightColorValues.None) {
                    text = $"=={text}==";
                }

                bool code = !string.IsNullOrEmpty(codeFont) && string.Equals(run.FontFamily, codeFont, StringComparison.OrdinalIgnoreCase);
                if (code) {
                    text = $"`{text}`";
                }

                if (run.IsHyperLink && run.Hyperlink != null && run.Hyperlink.Uri != null) {
                    text = $"[{text}]({run.Hyperlink.Uri})";
                }

                sb.Append(text);
            }

            return sb.ToString();
        }

        private string RenderFootnote(WordFootNote footNote, WordToMarkdownOptions options) {
            var paragraphs = footNote.Paragraphs;
            if (paragraphs == null || paragraphs.Count < 2) {
                return string.Empty;
            }

            var sb = new StringBuilder();
            for (int i = 1; i < paragraphs.Count; i++) {
                if (i > 1) {
                    sb.Append(' ');
                }
                sb.Append(RenderRuns(paragraphs[i], options));
            }
            return sb.ToString();
        }

        private string RenderImage(WordImage image, WordToMarkdownOptions options) {
            if (image == null) {
                return string.Empty;
            }

            string alt = image.Description ?? string.Empty;

            if (options.ImageExportMode == ImageExportMode.File) {
                string directory = options.ImageDirectory ?? Directory.GetCurrentDirectory();
                Directory.CreateDirectory(directory);
                string extension = Path.GetExtension(image.FilePath);
                if (string.IsNullOrEmpty(extension)) {
                    extension = ".png";
                }
                string fileName = string.IsNullOrEmpty(image.FileName)
                    ? Guid.NewGuid().ToString("N") + extension
                    : image.FileName!;
                string targetPath = Path.Combine(directory, fileName);

                if (!string.IsNullOrEmpty(image.FilePath) && File.Exists(image.FilePath)) {
                    File.Copy(image.FilePath, targetPath, true);
                } else {
                    File.WriteAllBytes(targetPath, image.GetBytes());
                }

                return $"![{alt}]({fileName})";
            } else {
                byte[] bytes = image.GetBytes();
                string extension = Path.GetExtension(image.FilePath);
                string mime = extension switch {
                    ".jpg" => "image/jpeg",
                    ".jpeg" => "image/jpeg",
                    ".gif" => "image/gif",
                    ".bmp" => "image/bmp",
                    _ => "image/png"
                };
                string base64 = System.Convert.ToBase64String(bytes);
                return $"![{alt}](data:{mime};base64,{base64})";
            }
        }
    }
}
