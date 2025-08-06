using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html.Helpers;
using System;
using System.Collections.Generic;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private struct TextFormatting {
            public bool Bold;
            public bool Italic;
            public bool Underline;

            public TextFormatting(bool bold = false, bool italic = false, bool underline = false) {
                Bold = bold;
                Italic = italic;
                Underline = underline;
            }
        }

        private static void ApplyParagraphStyleFromCss(WordParagraph paragraph, IElement element) {
            var style = CssStyleMapper.MapParagraphStyle(element.GetAttribute("style"));
            if (style.HasValue) {
                paragraph.Style = style.Value;
            }
        }

        private static readonly System.Text.RegularExpressions.Regex _urlRegex = new(@"((?:https?|ftp)://[^\s]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        private static void AddTextRun(WordParagraph paragraph, string text, TextFormatting formatting, HtmlToWordOptions options) {
            int lastIndex = 0;
            foreach (System.Text.RegularExpressions.Match match in _urlRegex.Matches(text)) {
                if (match.Index > lastIndex) {
                    var segment = text.Substring(lastIndex, match.Index - lastIndex);
                    var run = paragraph.AddFormattedText(segment, formatting.Bold, formatting.Italic, formatting.Underline ? UnderlineValues.Single : null);
                    if (!string.IsNullOrEmpty(options.FontFamily)) {
                        run.SetFontFamily(options.FontFamily);
                    }
                }
                var linkRun = paragraph.AddHyperLink(match.Value, new Uri(match.Value));
                ApplyFormatting(linkRun, formatting, options);
                lastIndex = match.Index + match.Length;
            }
            if (lastIndex < text.Length) {
                var run = paragraph.AddFormattedText(text.Substring(lastIndex), formatting.Bold, formatting.Italic, formatting.Underline ? UnderlineValues.Single : null);
                if (!string.IsNullOrEmpty(options.FontFamily)) {
                    run.SetFontFamily(options.FontFamily);
                }
            }
        }

        private static void ApplyFormatting(WordParagraph run, TextFormatting formatting, HtmlToWordOptions options) {
            if (formatting.Bold) run.SetBold();
            if (formatting.Italic) run.SetItalic();
            if (formatting.Underline) run.SetUnderline(UnderlineValues.Single);
            if (!string.IsNullOrEmpty(options.FontFamily)) run.SetFontFamily(options.FontFamily);
        }
    }
}