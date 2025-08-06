using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html.Helpers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Color = SixLabors.ImageSharp.Color;
using SixLabors.ImageSharp.PixelFormats;

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
            if (value.StartsWith("rgb", StringComparison.OrdinalIgnoreCase)) {
                int start = value.IndexOf('(');
                int end = value.IndexOf(')');
                if (start >= 0 && end > start) {
                    var parts = value.Substring(start + 1, end - start - 1).Split(',');
                    if (parts.Length >= 3 &&
                        byte.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte r) &&
                        byte.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte g) &&
                        byte.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte b)) {
                        var color = new Color(new Rgb24(r, g, b));
                        return color.ToHexColor();
                    }
                }
                return null;
            }
            try {
                var parsed = Color.Parse(value);
                return parsed.ToHexColor();
            } catch {
                if (!value.StartsWith("#", StringComparison.Ordinal)) {
                    try {
                        var parsed = Color.Parse("#" + value);
                        return parsed.ToHexColor();
                    } catch {
                        return null;
                    }
                }
                return null;
            }
        }

        private static readonly Dictionary<HighlightColorValues, Color> _highlightColors = new() {
            { HighlightColorValues.Yellow, Color.Yellow },
            { HighlightColorValues.Green, Color.Lime },
            { HighlightColorValues.Cyan, Color.Cyan },
            { HighlightColorValues.Magenta, Color.Magenta },
            { HighlightColorValues.Blue, Color.Blue },
            { HighlightColorValues.Red, Color.Red },
            { HighlightColorValues.DarkBlue, Color.DarkBlue },
            { HighlightColorValues.DarkCyan, Color.DarkCyan },
            { HighlightColorValues.DarkGreen, Color.DarkGreen },
            { HighlightColorValues.DarkMagenta, Color.DarkMagenta },
            { HighlightColorValues.DarkRed, Color.DarkRed },
            { HighlightColorValues.DarkYellow, Color.Parse("#808000") },
            { HighlightColorValues.DarkGray, Color.DarkGray },
            { HighlightColorValues.LightGray, Color.LightGray },
            { HighlightColorValues.Black, Color.Black },
            { HighlightColorValues.White, Color.White }
        };

        private static HighlightColorValues? MapColorToHighlight(string hex) {
            try {
                var target = Color.Parse("#" + hex);
                var targetRgb = target.ToPixel<Rgb24>();
                HighlightColorValues? best = null;
                int bestDistance = int.MaxValue;
                foreach (var pair in _highlightColors) {
                    var rgb = pair.Value.ToPixel<Rgb24>();
                    int distance = (rgb.R - targetRgb.R) * (rgb.R - targetRgb.R) +
                                   (rgb.G - targetRgb.G) * (rgb.G - targetRgb.G) +
                                   (rgb.B - targetRgb.B) * (rgb.B - targetRgb.B);
                    if (distance < bestDistance) {
                        bestDistance = distance;
                        best = pair.Key;
                    }
                }
                return best;
            } catch {
                return null;
            }
        }
    }
}