using AngleSharp.Dom;
using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Css.Values;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html.Helpers;
using System;
using System.Collections.Generic;
using System.Globalization;
using Color = SixLabors.ImageSharp.Color;
using SixLabors.ImageSharp.PixelFormats;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private struct TextFormatting {
            public bool Bold;
            public bool Italic;
            public bool Underline;
            public bool Strike;
            public bool Superscript;
            public bool Subscript;
            public string? ColorHex;
            public string? FontFamily;
            public int? FontSize;
            public HighlightColorValues? Highlight;

            public TextFormatting(bool bold = false, bool italic = false, bool underline = false, string? colorHex = null, string? fontFamily = null, int? fontSize = null, bool superscript = false, bool subscript = false, bool strike = false, HighlightColorValues? highlight = null) {
                Bold = bold;
                Italic = italic;
                Underline = underline;
                Strike = strike;
                Superscript = superscript;
                Subscript = subscript;
                ColorHex = colorHex;
                FontFamily = fontFamily;
                FontSize = fontSize;
                Highlight = highlight;
            }
        }

        private static readonly DefaultRenderDevice _renderDevice = new() { FontSize = 16 };
        private static readonly CssParser _inlineParser = new();

        private static bool TryParseFontSize(string? text, out int size) {
            size = 0;
            if (string.IsNullOrWhiteSpace(text)) {
                return false;
            }
            text = text.Trim().ToLowerInvariant();
            if (text.EndsWith("pt") && double.TryParse(text[..^2], out double pt)) {
                size = (int)Math.Round(pt);
                return size > 0;
            }
            if (text.EndsWith("px") && double.TryParse(text[..^2], out double px)) {
                size = (int)Math.Round(px);
                return size > 0;
            }
            return false;
        }

        private static bool TryConvertToTwip(ICssValue? value, out int twips) {
            twips = 0;
            if (value is CssLengthValue length) {
                try {
                    double px = length.ToPixel(_renderDevice);
                    twips = (int)Math.Round(px * 15);
                    return twips > 0;
                } catch { }
            }
            return false;
        }

        private static void ApplyParagraphStyleFromCss(WordParagraph paragraph, IElement element) {
            string? styleAttribute = element.GetAttribute("style");
            var style = CssStyleMapper.MapParagraphStyle(styleAttribute);
            if (style.HasValue) {
                paragraph.Style = style.Value;
            }

            var styleText = element.GetAttribute("style") ?? string.Empty;
            var declaration = _inlineParser.ParseDeclaration(styleText);
            if (declaration.Length == 0) {
                return;
            }

            int? marginLeft = null, marginRight = null, marginTop = null, marginBottom = null;
            int? paddingLeft = null, paddingRight = null, paddingTop = null, paddingBottom = null;
            JustificationValues? alignment = null;

            var colorVal = NormalizeColor(declaration.GetPropertyValue("color"));
            if (colorVal != null) {
                paragraph.SetColorHex(colorVal);
            }

            var bgColorVal = NormalizeColor(declaration.GetPropertyValue("background-color"));
            if (bgColorVal != null) {
                var highlight = MapColorToHighlight(bgColorVal);
                if (highlight.HasValue) {
                    paragraph.SetHighlight(highlight.Value);
                }
            }

            if (TryParseFontSize(declaration.GetPropertyValue("font-size"), out int fontSize)) {
                paragraph.SetFontSize(fontSize);
            }

            var align = declaration.GetPropertyValue("text-align")?.Trim();
            if (!string.IsNullOrEmpty(align)) {
                alignment = align.ToLowerInvariant() switch {
                    "center" => JustificationValues.Center,
                    "right" => JustificationValues.Right,
                    "justify" => JustificationValues.Both,
                    "left" => JustificationValues.Left,
                    _ => alignment
                };
            }

            if (TryConvertToTwip(declaration.GetProperty("margin-left")?.RawValue, out int ml)) marginLeft = ml;
            if (TryConvertToTwip(declaration.GetProperty("margin-right")?.RawValue, out int mr)) marginRight = mr;
            if (TryConvertToTwip(declaration.GetProperty("margin-top")?.RawValue, out int mt)) marginTop = mt;
            if (TryConvertToTwip(declaration.GetProperty("margin-bottom")?.RawValue, out int mb)) marginBottom = mb;
            if (TryConvertToTwip(declaration.GetProperty("padding-left")?.RawValue, out int pl)) paddingLeft = pl;
            if (TryConvertToTwip(declaration.GetProperty("padding-right")?.RawValue, out int pr)) paddingRight = pr;
            if (TryConvertToTwip(declaration.GetProperty("padding-top")?.RawValue, out int pt)) paddingTop = pt;
            if (TryConvertToTwip(declaration.GetProperty("padding-bottom")?.RawValue, out int pb)) paddingBottom = pb;

            if (alignment.HasValue) {
                paragraph.ParagraphAlignment = alignment;
            }
            int before = (marginTop ?? 0) + (paddingTop ?? 0);
            if (before > 0) {
                paragraph.LineSpacingBefore = before;
            }
            int after = (marginBottom ?? 0) + (paddingBottom ?? 0);
            if (after > 0) {
                paragraph.LineSpacingAfter = after;
            }
            int left = (marginLeft ?? 0) + (paddingLeft ?? 0);
            if (left > 0) {
                paragraph.IndentationBefore = left;
            }
            int right = (marginRight ?? 0) + (paddingRight ?? 0);
            if (right > 0) {
                paragraph.IndentationAfter = right;
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
            if (formatting.Strike) run.SetStrike();
            if (formatting.Superscript) run.SetSuperScript();
            if (formatting.Subscript) run.SetSubScript();
            if (!string.IsNullOrEmpty(formatting.ColorHex)) run.SetColorHex(formatting.ColorHex);
            if (formatting.Highlight.HasValue) run.SetHighlight(formatting.Highlight.Value);
            if (formatting.FontSize.HasValue) run.SetFontSize(formatting.FontSize.Value);
            if (!string.IsNullOrEmpty(formatting.FontFamily)) {
                run.SetFontFamily(formatting.FontFamily);
            } else if (!string.IsNullOrEmpty(options.FontFamily)) {
                run.SetFontFamily(options.FontFamily);
            }
        }

        private static void ApplySpanStyles(IElement element, ref TextFormatting formatting) {
            var styleText = element.GetAttribute("style") ?? string.Empty;
            var declaration = _inlineParser.ParseDeclaration(styleText);
            if (declaration.Length == 0) {
                return;
            }

            var color = NormalizeColor(declaration.GetPropertyValue("color"));
            if (color != null) {
                formatting.ColorHex = color;
            }

            var family = declaration.GetPropertyValue("font-family");
            if (!string.IsNullOrWhiteSpace(family)) {
                formatting.FontFamily = family.Trim('"', '\'', ' ');
            }

            if (TryParseFontSize(declaration.GetPropertyValue("font-size"), out int size)) {
                formatting.FontSize = size;
            }

            var weight = declaration.GetPropertyValue("font-weight");
            if (!string.IsNullOrEmpty(weight)) {
                if (int.TryParse(weight, out int w)) {
                    formatting.Bold = w >= 600;
                } else if (string.Equals(weight, "bold", StringComparison.OrdinalIgnoreCase)) {
                    formatting.Bold = true;
                } else if (string.Equals(weight, "normal", StringComparison.OrdinalIgnoreCase)) {
                    formatting.Bold = false;
                }
            }

            var fontStyle = declaration.GetPropertyValue("font-style").ToLowerInvariant();
            if (fontStyle == "italic" || fontStyle == "oblique") {
                formatting.Italic = true;
            } else if (fontStyle == "normal") {
                formatting.Italic = false;
            }

            var va = declaration.GetPropertyValue("vertical-align").ToLowerInvariant();
            if (va == "super" || va == "sup") {
                formatting.Superscript = true;
                formatting.Subscript = false;
            } else if (va == "sub") {
                formatting.Subscript = true;
                formatting.Superscript = false;
            } else if (va == "baseline") {
                formatting.Superscript = false;
                formatting.Subscript = false;
            }

            var decoValue = declaration.GetPropertyValue("text-decoration");
            foreach (var deco in decoValue.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                switch (deco.Trim().ToLowerInvariant()) {
                    case "underline":
                        formatting.Underline = true;
                        break;
                    case "line-through":
                        formatting.Strike = true;
                        break;
                }
            }

            var bgColor = NormalizeColor(declaration.GetPropertyValue("background-color"));
            if (bgColor != null) {
                var highlight = MapColorToHighlight(bgColor);
                if (highlight.HasValue) {
                    formatting.Highlight = highlight.Value;
                }
            }
        }

        private static string MergeStyles(string parentStyle, string? childStyle) {
            var parser = new CssParser();
            var parent = parser.ParseDeclaration(parentStyle);
            var child = parser.ParseDeclaration(childStyle ?? string.Empty);
            foreach (var prop in parent) {
                if (string.IsNullOrEmpty(child.GetPropertyValue(prop.Name))) {
                    child.SetProperty(prop.Name, prop.Value);
                }
            }
            return child.CssText;
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