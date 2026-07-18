using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointTextBox {
        /// <summary>
        ///     Replaces the textbox content with rich text parsed from a small Markdown subset.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetMarkdown(string markdown) {
            if (markdown == null) {
                throw new ArgumentNullException(nameof(markdown));
            }

            TextBody textBody = EnsureTextBody();
            string[] discardedSoundIds = PowerPointEmbeddedSound
                .GetRelationshipIds(textBody);
            A.Paragraph? templateParagraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
            textBody.RemoveAllChildren<A.Paragraph>();
            IReadOnlyList<PowerPointParagraph> paragraphs = AppendMarkdown(
                markdown, templateParagraph);
            PowerPointEmbeddedSound.RemoveIfUnused(_slidePart,
                discardedSoundIds);
            return paragraphs;
        }

        /// <summary>
        ///     Appends rich text parsed from a small Markdown subset to the textbox.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddMarkdown(string markdown) {
            if (markdown == null) {
                throw new ArgumentNullException(nameof(markdown));
            }

            TextBody textBody = EnsureTextBody();
            A.Paragraph? templateParagraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
            return AppendMarkdown(markdown, templateParagraph);
        }

        private IReadOnlyList<PowerPointParagraph> AppendMarkdown(string markdown, A.Paragraph? templateParagraph) {
            TextBody textBody = EnsureTextBody();
            string[] lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var results = new List<PowerPointParagraph>();

            foreach (string rawLine in lines) {
                string line = rawLine.TrimEnd();
                if (line.Length == 0) {
                    continue;
                }

                MarkdownParagraphKind kind = GetMarkdownParagraphKind(line, out string text, out int headingLevel,
                    out int listStart);
                PowerPointParagraph paragraph = AppendParagraph(textBody, string.Empty, templateParagraph);
                ApplyMarkdownInline(paragraph, ParseMarkdownInline(text));
                ApplyMarkdownParagraphKind(paragraph, kind, headingLevel, listStart);
                results.Add(paragraph);
            }

            if (results.Count == 0 && textBody.Elements<A.Paragraph>().FirstOrDefault() == null) {
                textBody.Append(CreateEmptyParagraph(templateParagraph));
            }

            return results;
        }

        private static MarkdownParagraphKind GetMarkdownParagraphKind(string line, out string text, out int headingLevel,
            out int listStart) {
            text = line;
            headingLevel = 0;
            listStart = 1;

            int hashCount = CountLeading(line, '#');
            if (hashCount is >= 1 and <= 6 && line.Length > hashCount && line[hashCount] == ' ') {
                text = line.Substring(hashCount + 1).TrimStart();
                headingLevel = hashCount;
                return MarkdownParagraphKind.Heading;
            }

            string trimmed = line.TrimStart();
            if (trimmed.Length > 2 && (trimmed[0] == '-' || trimmed[0] == '*' || trimmed[0] == '+') &&
                trimmed[1] == ' ') {
                text = trimmed.Substring(2).TrimStart();
                return MarkdownParagraphKind.Bullet;
            }

            int dotIndex = trimmed.IndexOf('.');
            if (dotIndex > 0 && dotIndex < 8 &&
                dotIndex + 1 < trimmed.Length &&
                trimmed[dotIndex + 1] == ' ' &&
                int.TryParse(trimmed.Substring(0, dotIndex), out int parsedStart)) {
                text = trimmed.Substring(dotIndex + 2).TrimStart();
                listStart = parsedStart;
                return MarkdownParagraphKind.Numbered;
            }

            return MarkdownParagraphKind.Normal;
        }

        private static int CountLeading(string value, char character) {
            int count = 0;
            while (count < value.Length && value[count] == character) {
                count++;
            }

            return count;
        }

        private static void ApplyMarkdownParagraphKind(PowerPointParagraph paragraph, MarkdownParagraphKind kind,
            int headingLevel, int listStart) {
            if (kind == MarkdownParagraphKind.Bullet) {
                paragraph.SetBullet();
                return;
            }

            if (kind == MarkdownParagraphKind.Numbered) {
                paragraph.SetNumbered(listStart);
                return;
            }

            if (kind == MarkdownParagraphKind.Heading) {
                int fontSize = headingLevel switch {
                    1 => 28,
                    2 => 24,
                    3 => 20,
                    _ => 18
                };

                foreach (PowerPointTextRun run in paragraph.Runs) {
                    run.Bold = true;
                    run.FontSize = fontSize;
                }
            }
        }

        private static void ApplyMarkdownInline(PowerPointParagraph paragraph, IReadOnlyList<MarkdownInlineRun> runs) {
            if (runs.Count == 0) {
                paragraph.Text = string.Empty;
                return;
            }

            bool first = true;
            foreach (MarkdownInlineRun inlineRun in runs) {
                PowerPointTextRun run;
                if (first) {
                    paragraph.Text = inlineRun.Text;
                    run = paragraph.Runs[0];
                    first = false;
                } else {
                    run = paragraph.AddRun(inlineRun.Text);
                }

                ApplyMarkdownInlineStyle(run, inlineRun);
            }
        }

        private static void ApplyMarkdownInlineStyle(PowerPointTextRun run, MarkdownInlineRun inlineRun) {
            if (inlineRun.Bold) {
                run.Bold = true;
            }

            if (inlineRun.Italic) {
                run.Italic = true;
            }

            if (inlineRun.Strikethrough) {
                run.Strikethrough = true;
            }

            if (inlineRun.Code) {
                run.FontName = "Consolas";
                run.HighlightColor = "F2F2F2";
            }

            if (!string.IsNullOrWhiteSpace(inlineRun.Link)) {
                run.SetHyperlink(inlineRun.Link!);
                run.Underline = true;
                run.Color = "0563C1";
            }
        }

        private static IReadOnlyList<MarkdownInlineRun> ParseMarkdownInline(string text) {
            var runs = new List<MarkdownInlineRun>();
            int index = 0;
            while (index < text.Length) {
                if (TryParseDelimited(text, index, "**", out MarkdownInlineRun boldRun, bold: true)) {
                    runs.Add(boldRun);
                    index += boldRun.SourceLength;
                    continue;
                }

                if (TryParseDelimited(text, index, "__", out MarkdownInlineRun underscoreBoldRun, bold: true)) {
                    runs.Add(underscoreBoldRun);
                    index += underscoreBoldRun.SourceLength;
                    continue;
                }

                if (TryParseDelimited(text, index, "~~", out MarkdownInlineRun strikeRun, strikethrough: true)) {
                    runs.Add(strikeRun);
                    index += strikeRun.SourceLength;
                    continue;
                }

                if (TryParseDelimited(text, index, "`", out MarkdownInlineRun codeRun, code: true)) {
                    runs.Add(codeRun);
                    index += codeRun.SourceLength;
                    continue;
                }

                if (TryParseLink(text, index, out MarkdownInlineRun linkRun)) {
                    runs.Add(linkRun);
                    index += linkRun.SourceLength;
                    continue;
                }

                if (TryParseDelimited(text, index, "*", out MarkdownInlineRun italicRun, italic: true)) {
                    runs.Add(italicRun);
                    index += italicRun.SourceLength;
                    continue;
                }

                if (TryParseDelimited(text, index, "_", out MarkdownInlineRun underscoreItalicRun, italic: true)) {
                    runs.Add(underscoreItalicRun);
                    index += underscoreItalicRun.SourceLength;
                    continue;
                }

                int next = FindNextMarkdownMarker(text, index + 1);
                string plain = text.Substring(index, next - index);
                runs.Add(new MarkdownInlineRun(plain, next - index));
                index = next;
            }

            return runs;
        }

        private static bool TryParseDelimited(string text, int index, string delimiter, out MarkdownInlineRun run,
            bool bold = false, bool italic = false, bool code = false, bool strikethrough = false) {
            run = default;
            if (!text.Substring(index).StartsWith(delimiter, StringComparison.Ordinal)) {
                return false;
            }

            int contentStart = index + delimiter.Length;
            int end = text.IndexOf(delimiter, contentStart, StringComparison.Ordinal);
            if (end <= contentStart) {
                return false;
            }

            string content = text.Substring(contentStart, end - contentStart);
            run = new MarkdownInlineRun(content, end + delimiter.Length - index, bold, italic, code, strikethrough);
            return true;
        }

        private static bool TryParseLink(string text, int index, out MarkdownInlineRun run) {
            run = default;
            if (text[index] != '[') {
                return false;
            }

            int labelEnd = text.IndexOf("](", index, StringComparison.Ordinal);
            if (labelEnd < 0) {
                return false;
            }

            int urlStart = labelEnd + 2;
            int urlEnd = text.IndexOf(')', urlStart);
            if (urlEnd <= urlStart) {
                return false;
            }

            string label = text.Substring(index + 1, labelEnd - index - 1);
            string url = text.Substring(urlStart, urlEnd - urlStart);
            run = new MarkdownInlineRun(label, urlEnd + 1 - index, link: url);
            return true;
        }

        private static int FindNextMarkdownMarker(string text, int start) {
            int next = text.Length;
            foreach (char marker in new[] { '*', '_', '`', '[', '~' }) {
                int markerIndex = text.IndexOf(marker, start);
                if (markerIndex >= 0 && markerIndex < next) {
                    next = markerIndex;
                }
            }

            return next;
        }

        private enum MarkdownParagraphKind {
            Normal,
            Heading,
            Bullet,
            Numbered
        }

        private readonly struct MarkdownInlineRun {
            public MarkdownInlineRun(string text, int sourceLength, bool bold = false, bool italic = false,
                bool code = false, bool strikethrough = false, string? link = null) {
                Text = text;
                SourceLength = sourceLength;
                Bold = bold;
                Italic = italic;
                Code = code;
                Strikethrough = strikethrough;
                Link = link;
            }

            public string Text { get; }

            public int SourceLength { get; }

            public bool Bold { get; }

            public bool Italic { get; }

            public bool Code { get; }

            public bool Strikethrough { get; }

            public string? Link { get; }
        }
    }
}
