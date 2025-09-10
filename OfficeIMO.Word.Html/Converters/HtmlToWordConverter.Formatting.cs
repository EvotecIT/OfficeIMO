using AngleSharp.Dom;
using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Css.Values;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Color = SixLabors.ImageSharp.Color;
using SixLabors.ImageSharp.PixelFormats;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private enum TextTransform {
            None,
            Uppercase,
            Lowercase,
            Capitalize,
        }

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
            public CapsStyle? Caps;
            public int? LetterSpacing;
            public TextTransform Transform;
            public WhiteSpaceMode? WhiteSpace;

            public TextFormatting(bool bold = false, bool italic = false, bool underline = false, string? colorHex = null, string? fontFamily = null, int? fontSize = null, bool superscript = false, bool subscript = false, bool strike = false, HighlightColorValues? highlight = null, int? letterSpacing = null, TextTransform transform = TextTransform.None, WhiteSpaceMode? whiteSpace = null) {
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
                Caps = null;
                LetterSpacing = letterSpacing;
                Transform = transform;
                WhiteSpace = whiteSpace;
            }
        }

        private static readonly DefaultRenderDevice _renderDevice = new() { FontSize = 16 };
        private static readonly CssParser _inlineParser = new();
        private static readonly Dictionary<string, int> _namedFontSizes = new(StringComparer.OrdinalIgnoreCase) {
            { "xx-small", 9 },
            { "x-small", 10 },
            { "small", 13 },
            { "medium", 16 },
            { "large", 18 },
            { "x-large", 24 },
            { "xx-large", 32 },
        };

        private static bool TryParseFontSize(string? text, out int size) {
            size = 0;
            if (string.IsNullOrWhiteSpace(text)) {
                return false;
            }
            text = (text ?? string.Empty).Trim().ToLowerInvariant();
            if (text.EndsWith("pt") && double.TryParse(text.Substring(0, text.Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out double pt)) {
                size = (int)Math.Round(pt);
                return size > 0;
            }
            if (text.EndsWith("px") && double.TryParse(text.Substring(0, text.Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out double px)) {
                size = (int)Math.Round(px);
                return size > 0;
            }
            if (_namedFontSizes.TryGetValue(text, out int named)) {
                size = named;
                return true;
            }
            if (text.EndsWith("%") && double.TryParse(text.Substring(0, text.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double percent)) {
                size = (int)Math.Round(_renderDevice.FontSize * (percent / 100d));
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

        private static bool TryConvertToTwipAllowNegative(ICssValue? value, out int twips) {
            twips = 0;
            if (value is CssLengthValue length) {
                try {
                    double px = length.ToPixel(_renderDevice);
                    twips = (int)Math.Round(px * 15);
                    return true;
                } catch { }
            }
            return false;
        }

        private static TextTransform? ParseTextTransform(string? value) =>
            value?.Trim().ToLowerInvariant() switch {
                "uppercase" => TextTransform.Uppercase,
                "lowercase" => TextTransform.Lowercase,
                "capitalize" => TextTransform.Capitalize,
                _ => null,
            };

        private static string ApplyTextTransform(string text, TextTransform? transform) {
            if (transform == null || transform == TextTransform.None || string.IsNullOrEmpty(text)) {
                return text;
            }
            return transform.Value switch {
                TextTransform.Uppercase => text.ToUpperInvariant(),
                TextTransform.Lowercase => text.ToLowerInvariant(),
                TextTransform.Capitalize => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(text.ToLowerInvariant()),
                _ => text,
            };
        }

        private static bool IsGenericFont(string family) =>
            string.Equals(family, "serif", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(family, "sans-serif", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(family, "monospace", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(family, "cursive", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(family, "fantasy", StringComparison.OrdinalIgnoreCase);

        private static string? ResolveFontFamily(string? family) {
            if (string.IsNullOrWhiteSpace(family)) {
                return null;
            }
            foreach (var part in (family ?? string.Empty).Split(',')) {
                var trimmed = part.Trim('"', '\'', ' ');
                if (string.IsNullOrEmpty(trimmed)) {
                    continue;
                }
                if (IsGenericFont(trimmed)) {
                    var resolved = FontResolver.Resolve(trimmed);
                    if (!string.IsNullOrEmpty(resolved)) {
                        return resolved;
                    }
                } else {
                    return trimmed;
                }
            }
            return null;
        }

        private static CssStyleMapper.CssProperties ApplyParagraphStyleFromCss(WordParagraph paragraph, IElement element) {
            string? styleAttribute = element.GetAttribute("style");
            var style = CssStyleMapper.MapParagraphStyle(styleAttribute);
            if (style.HasValue) {
                paragraph.Style = style.Value;
            }

            var parsed = CssStyleMapper.ParseStyles(styleAttribute);
            var declaration = _inlineParser.ParseDeclaration(styleAttribute ?? string.Empty);
            int? marginLeft = parsed.MarginLeft, marginRight = parsed.MarginRight, marginTop = parsed.MarginTop, marginBottom = parsed.MarginBottom;
            int? paddingLeft = null, paddingRight = null, paddingTop = null, paddingBottom = null;
            JustificationValues? alignment = null;

            var colorVal = NormalizeColor(declaration.GetPropertyValue("color"));
            if (colorVal != null) {
                paragraph.SetColorHex(colorVal);
            }

            if (!string.IsNullOrEmpty(parsed.BackgroundColor)) {
                var highlight = MapColorToHighlight(parsed.BackgroundColor);
                if (highlight.HasValue) {
                    paragraph.SetHighlight(highlight.Value);
                }
            }

            if (parsed.LineHeight.HasValue) {
                paragraph.LineSpacing = parsed.LineHeight.Value;
                if (parsed.LineHeightRule.HasValue) {
                    paragraph.LineSpacingRule = parsed.LineHeightRule.Value;
                }
            }

            if (TryParseFontSize(declaration.GetPropertyValue("font-size"), out int fontSize)) {
                paragraph.SetFontSize(fontSize);
            }

            if (TryConvertToTwipAllowNegative(declaration.GetProperty("letter-spacing")?.RawValue, out int lsParagraph)) {
                paragraph.SetSpacing(lsParagraph);
            }

            var paragraphTransform = ParseTextTransform(declaration.GetPropertyValue("text-transform"));
            if (paragraphTransform.HasValue && !string.IsNullOrEmpty(paragraph.Text)) {
                paragraph.SetText(ApplyTextTransform(paragraph.Text, paragraphTransform.Value));
            }

            var align = declaration.GetPropertyValue("text-align")?.Trim();
            if (!string.IsNullOrEmpty(align)) {
                alignment = align!.ToLowerInvariant() switch {
                    "center" => JustificationValues.Center,
                    "right" => JustificationValues.Right,
                    "justify" => JustificationValues.Both,
                    "left" => JustificationValues.Left,
                    _ => alignment
                };
            }

            var floatVal = declaration.GetPropertyValue("float")?.Trim();
            if (floatVal != null) {
                var f = floatVal.ToLowerInvariant();
                alignment = f switch {
                    "left" => JustificationValues.Left,
                    "right" => JustificationValues.Right,
                    _ => alignment
                };
            }

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
            return parsed;
        }

        private static readonly Regex _urlRegex = new(@"((?:https?|ftp)://[^\s]+)", RegexOptions.IgnoreCase);
        private static readonly Regex _collapseWhitespaceRegex = new(@"\s+", RegexOptions.Compiled);

        private static void AddTextRun(WordParagraph paragraph, string text, TextFormatting formatting, HtmlToWordOptions options) {
            text = ApplyWhiteSpace(text, formatting.WhiteSpace);
            int lastIndex = 0;
            foreach (Match match in _urlRegex.Matches(text)) {
                if (match.Index > lastIndex) {
                    var segment = text.Substring(lastIndex, match.Index - lastIndex);
                    segment = ApplyTextTransform(segment, formatting.Transform);
                    var run = paragraph.AddFormattedText(segment, formatting.Bold, formatting.Italic, formatting.Underline ? UnderlineValues.Single : null);
                    ApplyFormatting(run, formatting, options);
                }
                var display = ApplyTextTransform(match.Value, formatting.Transform);
                var linkRun = paragraph.AddHyperLink(display, new Uri(match.Value));
                ApplyFormatting(linkRun, formatting, options);
                lastIndex = match.Index + match.Length;
            }
            if (lastIndex < text.Length) {
                var segment = text.Substring(lastIndex);
                segment = ApplyTextTransform(segment, formatting.Transform);
                var run = paragraph.AddFormattedText(segment, formatting.Bold, formatting.Italic, formatting.Underline ? UnderlineValues.Single : null);
                ApplyFormatting(run, formatting, options);
            }
        }

        private static string ApplyWhiteSpace(string text, WhiteSpaceMode? mode) {
            if (!mode.HasValue) {
                return text;
            }
            return mode.Value switch {
                WhiteSpaceMode.Normal => CollapseWhiteSpace(text, false),
                WhiteSpaceMode.NoWrap => CollapseWhiteSpace(text, true),
                WhiteSpaceMode.Pre => text.Replace(" ", " "),
                WhiteSpaceMode.PreWrap => text,
                _ => text,
            };
        }

        private static string CollapseWhiteSpace(string text, bool noWrap) {
            var collapsed = _collapseWhitespaceRegex.Replace(text, " ");
            if (noWrap) {
                collapsed = collapsed.Replace(" ", " ");
            }
            return collapsed;
        }

        private static void ApplyFormatting(WordParagraph run, TextFormatting formatting, HtmlToWordOptions options) {
            if (formatting.Bold) run.SetBold();
            if (formatting.Italic) run.SetItalic();
            if (formatting.Underline) run.SetUnderline(UnderlineValues.Single);
            if (formatting.Strike) run.SetStrike();
            if (formatting.Superscript) run.SetSuperScript();
            if (formatting.Subscript) run.SetSubScript();
            if (!string.IsNullOrEmpty(formatting.ColorHex)) run.SetColorHex(formatting.ColorHex!);
            if (formatting.Highlight.HasValue) run.SetHighlight(formatting.Highlight.Value);
            if (formatting.FontSize.HasValue) run.SetFontSize(formatting.FontSize.Value);
            if (formatting.Caps.HasValue) run.SetCapsStyle(formatting.Caps.Value);
            if (formatting.LetterSpacing.HasValue) run.SetSpacing(formatting.LetterSpacing.Value);
            if (!string.IsNullOrEmpty(formatting.FontFamily)) {
                var font = ResolveFontFamily(formatting.FontFamily);
                if (!string.IsNullOrEmpty(font)) {
                    run.SetFontFamily(font!);
                }
            } else if (!string.IsNullOrEmpty(options.FontFamily)) {
                var font = ResolveFontFamily(options.FontFamily);
                if (!string.IsNullOrEmpty(font)) {
                    run.SetFontFamily(font!);
                }
            }
        }

        private static void ApplySpanStyles(IElement element, ref TextFormatting formatting) {
            var styleText = element.GetAttribute("style") ?? string.Empty;
            var parsed = CssStyleMapper.ParseStyles(styleText);
            var declaration = _inlineParser.ParseDeclaration(styleText);
            if (declaration.Length == 0 && parsed.BackgroundColor == null && !parsed.Underline && !parsed.Strike) {
                return;
            }

            var color = NormalizeColor(declaration.GetPropertyValue("color"));
            if (color != null) {
                formatting.ColorHex = color;
            }

            var family = declaration.GetPropertyValue("font-family");
            if (!string.IsNullOrWhiteSpace(family)) {
                var resolved = ResolveFontFamily(family);
                if (!string.IsNullOrEmpty(resolved)) {
                    formatting.FontFamily = resolved;
                }
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

            var fontVariant = declaration.GetPropertyValue("font-variant").ToLowerInvariant();
            if (fontVariant == "small-caps") {
                formatting.Caps = CapsStyle.SmallCaps;
            } else if (fontVariant == "normal") {
                formatting.Caps = CapsStyle.None;
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

            if (TryConvertToTwipAllowNegative(declaration.GetProperty("letter-spacing")?.RawValue, out int ls)) {
                formatting.LetterSpacing = ls;
            }

            var transform = ParseTextTransform(declaration.GetPropertyValue("text-transform"));
            if (transform.HasValue) {
                formatting.Transform = transform.Value;
            }

            if (parsed.Underline) {
                formatting.Underline = true;
            }
            if (parsed.Strike) {
                formatting.Strike = true;
            }
            if (!string.IsNullOrEmpty(parsed.BackgroundColor)) {
                var highlight = MapColorToHighlight(parsed.BackgroundColor);
                if (highlight.HasValue) {
                    formatting.Highlight = highlight.Value;
                }
            }
            if (parsed.WhiteSpace.HasValue) {
                formatting.WhiteSpace = parsed.WhiteSpace.Value;
            }
        }

        private static bool TryParseHtmlFontSize(string? value, out int size) {
            size = 0;
            if (TryParseFontSize(value, out size)) {
                return true;
            }
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numeric)) {
                size = numeric switch {
                    1 => 8,
                    2 => 10,
                    3 => 12,
                    4 => 14,
                    5 => 18,
                    6 => 24,
                    7 => 36,
                    _ => 0
                };
                return size > 0;
            }
            return false;
        }

        private static void ApplyFontStyles(IElement element, ref TextFormatting formatting) {
            ApplySpanStyles(element, ref formatting);

            var colorAttr = NormalizeColor(element.GetAttribute("color"));
            if (colorAttr != null) {
                formatting.ColorHex = colorAttr;
            }

            if (TryParseHtmlFontSize(element.GetAttribute("size"), out int size)) {
                formatting.FontSize = size;
            }
        }

        private static string MergeStyles(string? parentStyle, string? childStyle) {
            var parser = new CssParser();
            var parent = parser.ParseDeclaration(parentStyle ?? string.Empty);
            var child = parser.ParseDeclaration(childStyle ?? string.Empty);
            foreach (var prop in parent) {
                if (string.IsNullOrEmpty(child.GetPropertyValue(prop.Name))) {
                    child.SetProperty(prop.Name, prop.Value);
                }
            }
            return child.CssText;
        }


        private static string? NormalizeColor(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }
            string v = value!.Trim();
            if (v.StartsWith("rgb", StringComparison.OrdinalIgnoreCase)) {
                int start = v.IndexOf('(');
                int end = v.IndexOf(')');
                if (start >= 0 && end > start) {
                    var parts = v.Substring(start + 1, end - start - 1).Split(',');
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
                var parsed = Color.Parse(v);
                return parsed.ToHexColor();
            } catch {
                if (!v.StartsWith("#", StringComparison.Ordinal)) {
                    try {
                        var parsed = Color.Parse("#" + v);
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

        private static HighlightColorValues? MapColorToHighlight(string? hex) {
            if (string.IsNullOrEmpty(hex)) {
                return null;
            }
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
