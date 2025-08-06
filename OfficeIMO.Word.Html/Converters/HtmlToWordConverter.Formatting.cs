using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html.Helpers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private struct TextFormatting {
            public bool Bold;
            public bool Italic;
            public bool Underline;
            public string? ColorHex;
            public string? FontFamily;
            public int? FontSize;

            public TextFormatting(bool bold = false, bool italic = false, bool underline = false, string? colorHex = null, string? fontFamily = null, int? fontSize = null) {
                Bold = bold;
                Italic = italic;
                Underline = underline;
                ColorHex = colorHex;
                FontFamily = fontFamily;
                FontSize = fontSize;
            }
        }

        private static void ApplyParagraphStyleFromCss(WordParagraph paragraph, IElement element) {
            string? styleAttribute = element.GetAttribute("style");
            var style = CssStyleMapper.MapParagraphStyle(styleAttribute);
            if (style.HasValue) {
                paragraph.Style = style.Value;
            }

            if (string.IsNullOrWhiteSpace(styleAttribute)) {
                return;
            }

            foreach (var part in styleAttribute.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                var pieces = part.Split(new[] { ':' }, 2);
                if (pieces.Length != 2) {
                    continue;
                }
                var name = pieces[0].Trim().ToLowerInvariant();
                var value = pieces[1].Trim();
                switch (name) {
                    case "color":
                        var color = NormalizeColor(value);
                        if (color != null) {
                            paragraph.SetColorHex(color);
                        }
                        break;
                    case "background-color":
                        var bgColor = NormalizeColor(value);
                        if (bgColor != null) {
                            var highlight = MapColorToHighlight(bgColor);
                            if (highlight.HasValue) {
                                paragraph.SetHighlight(highlight.Value);
                            }
                        }
                        break;
                    case "font-size":
                        if (TryParseFontSize(value, out int size)) {
                            paragraph.SetFontSize(size);
                        }
                        break;
                }
            }
        }

        private static readonly System.Text.RegularExpressions.Regex _urlRegex = new(@"((?:https?|ftp)://[^\s]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        private static void AddTextRun(WordParagraph paragraph, string text, TextFormatting formatting, HtmlToWordOptions options) {
            int lastIndex = 0;
            foreach (System.Text.RegularExpressions.Match match in _urlRegex.Matches(text)) {
                if (match.Index > lastIndex) {
                    var segment = text.Substring(lastIndex, match.Index - lastIndex);
                    var run = paragraph.AddFormattedText(segment, formatting.Bold, formatting.Italic, formatting.Underline ? UnderlineValues.Single : null);
                    ApplyFormatting(run, formatting, options);
                }
                var linkRun = paragraph.AddHyperLink(match.Value, new Uri(match.Value));
                ApplyFormatting(linkRun, formatting, options);
                lastIndex = match.Index + match.Length;
            }
            if (lastIndex < text.Length) {
                var run = paragraph.AddFormattedText(text.Substring(lastIndex), formatting.Bold, formatting.Italic, formatting.Underline ? UnderlineValues.Single : null);
                ApplyFormatting(run, formatting, options);
            }
        }

        private static void ApplyFormatting(WordParagraph run, TextFormatting formatting, HtmlToWordOptions options) {
            if (formatting.Bold) run.SetBold();
            if (formatting.Italic) run.SetItalic();
            if (formatting.Underline) run.SetUnderline(UnderlineValues.Single);
            if (!string.IsNullOrEmpty(formatting.ColorHex)) run.SetColorHex(formatting.ColorHex);
            if (formatting.FontSize.HasValue) run.SetFontSize(formatting.FontSize.Value);
            if (!string.IsNullOrEmpty(formatting.FontFamily)) {
                run.SetFontFamily(formatting.FontFamily);
            } else if (!string.IsNullOrEmpty(options.FontFamily)) {
                run.SetFontFamily(options.FontFamily);
            }
        }

        private static void ApplySpanStyles(IElement element, ref TextFormatting formatting) {
            var style = element.GetAttribute("style");
            if (string.IsNullOrWhiteSpace(style)) {
                return;
            }

            foreach (var part in style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                var pieces = part.Split(new[] { ':' }, 2);
                if (pieces.Length != 2) {
                    continue;
                }
                var name = pieces[0].Trim().ToLowerInvariant();
                var value = pieces[1].Trim();
                switch (name) {
                    case "color":
                        var color = NormalizeColor(value);
                        if (color != null) {
                            formatting.ColorHex = color;
                        }
                        break;
                    case "font-family":
                        formatting.FontFamily = value.Trim('"', '\'', ' ');
                        break;
                    case "font-size":
                        if (TryParseFontSize(value, out int size)) {
                            formatting.FontSize = size;
                        }
                        break;
                }
            }
        }

        private static bool TryParseFontSize(string value, out int size) {
            size = 0;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }
            value = value.Trim().ToLowerInvariant();
            string number = new(value.Where(c => char.IsDigit(c) || c == '.').ToArray());
            if (!double.TryParse(number, NumberStyles.Number, CultureInfo.InvariantCulture, out double result)) {
                return false;
            }
            if (value.EndsWith("em", StringComparison.Ordinal)) {
                result *= 16;
            }
            size = (int)Math.Round(result);
            return size > 0;
        }

        private static string? NormalizeColor(string value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }
            value = value.Trim();
            if (value.StartsWith("#", StringComparison.Ordinal)) {
                var hex = value.Substring(1);
                if (hex.Length == 3) {
                    hex = string.Concat(hex.Select(c => new string(c, 2)));
                }
                return hex.ToLowerInvariant();
            }
            if (value.StartsWith("rgb", StringComparison.OrdinalIgnoreCase)) {
                int start = value.IndexOf('(');
                int end = value.IndexOf(')');
                if (start >= 0 && end > start) {
                    var parts = value.Substring(start + 1, end - start - 1).Split(',');
                    if (parts.Length >= 3 &&
                        byte.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte r) &&
                        byte.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte g) &&
                        byte.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte b)) {
                        var hex = string.Concat(r.ToString("X2"), g.ToString("X2"), b.ToString("X2"));
                        return hex.ToLowerInvariant();
                    }
                }
            }
            return null;
        }

        private static HighlightColorValues? MapColorToHighlight(string hex) {
            return hex.ToLowerInvariant() switch {
                "ffff00" => HighlightColorValues.Yellow,
                "00ff00" => HighlightColorValues.Green,
                "00ffff" => HighlightColorValues.Cyan,
                "ff00ff" => HighlightColorValues.Magenta,
                "0000ff" => HighlightColorValues.Blue,
                "ff0000" => HighlightColorValues.Red,
                "00008b" => HighlightColorValues.DarkBlue,
                "008b8b" => HighlightColorValues.DarkCyan,
                "006400" => HighlightColorValues.DarkGreen,
                "8b008b" => HighlightColorValues.DarkMagenta,
                "8b0000" => HighlightColorValues.DarkRed,
                "808000" => HighlightColorValues.DarkYellow,
                "a9a9a9" => HighlightColorValues.DarkGray,
                "d3d3d3" => HighlightColorValues.LightGray,
                "000000" => HighlightColorValues.Black,
                _ => null,
            };
        }
    }
}