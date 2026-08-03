using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static bool TryCreateRubyNode(IDocument htmlDoc, WordParagraph run, WordToHtmlOptions options, out INode node) {
            node = htmlDoc.CreateTextNode(string.Empty);
            var ruby = run._run?.Elements<Ruby>().FirstOrDefault();
            if (ruby == null) {
                return false;
            }

            var baseText = ruby.RubyBase?.InnerText ?? string.Empty;
            if (string.IsNullOrEmpty(baseText)) {
                return false;
            }

            var rubyText = ruby.RubyContent?.InnerText ?? string.Empty;
            if (string.IsNullOrEmpty(rubyText)) {
                node = htmlDoc.CreateTextNode(baseText);
                return true;
            }

            var rubyElement = CreateOutputElement(htmlDoc, "ruby");
            var baseElement = CreateOutputElement(htmlDoc, "rb");
            AppendRubyRuns(htmlDoc, baseElement, ruby.RubyBase, options);
            rubyElement.AppendChild(baseElement);

            var annotationElement = CreateOutputElement(htmlDoc, "rt");
            AppendRubyRuns(htmlDoc, annotationElement, ruby.RubyContent, options);
            rubyElement.AppendChild(annotationElement);

            node = rubyElement;
            return true;
        }

        private static void AppendRubyRuns(
            IDocument htmlDoc,
            IElement target,
            OpenXmlCompositeElement? source,
            WordToHtmlOptions options) {
            foreach (Run run in source?.Elements<Run>() ?? Enumerable.Empty<Run>()) {
                INode node = htmlDoc.CreateTextNode(run.InnerText);
                RunProperties? properties = run.RunProperties;
                if (properties == null) {
                    target.AppendChild(node);
                    continue;
                }
                if (IsEnabled(properties.Bold)) node = WrapRubyNode(htmlDoc, "strong", node);
                if (IsEnabled(properties.Italic)) node = WrapRubyNode(htmlDoc, "em", node);
                if (IsEnabled(properties.Strike)) node = WrapRubyNode(htmlDoc, "s", node);
                if (properties.Underline != null && properties.Underline.Val?.Value != UnderlineValues.None) {
                    node = WrapRubyNode(htmlDoc, "u", node);
                }
                if (properties.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Superscript) {
                    node = WrapRubyNode(htmlDoc, "sup", node);
                } else if (properties.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Subscript) {
                    node = WrapRubyNode(htmlDoc, "sub", node);
                }

                bool isMarked = string.Equals(
                    properties.RunStyle?.Val?.Value,
                    HtmlSemanticStyleIds.MarkedText,
                    StringComparison.OrdinalIgnoreCase);
                if (isMarked) node = WrapRubyNode(htmlDoc, "mark", node);

                var styles = new List<string>();
                string? color = NormalizeSixDigitHexColor(properties.Color?.Val?.Value);
                if (options.IncludeRunColorStyles && color != null) styles.Add($"color:#{color}");
                if (int.TryParse(properties.FontSize?.Val?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int halfPoints)) {
                    styles.Add($"font-size:{(halfPoints / 2D).ToString("0.###", CultureInfo.InvariantCulture)}pt");
                }
                string? font = properties.RunFonts?.Ascii?.Value ?? properties.RunFonts?.HighAnsi?.Value;
                if (options.IncludeFontStyles && !string.IsNullOrWhiteSpace(font)) styles.Add($"font-family:{QuoteCssString(font!)}");
                if (IsEnabled(properties.SmallCaps)) styles.Add("font-variant:small-caps");
                if (IsEnabled(properties.Caps)) styles.Add("text-transform:uppercase");
                if (properties.Spacing?.Val?.Value is int twentiethPoints) {
                    styles.Add($"letter-spacing:{(twentiethPoints / 20D).ToString("0.###", CultureInfo.InvariantCulture)}pt");
                }
                if (options.IncludeRunHighlightStyles) {
                    string? shading = NormalizeSixDigitHexColor(properties.Shading?.Fill?.Value);
                    string? highlight = GetHighlightCss(properties.Highlight?.Val?.Value);
                    if (!string.IsNullOrEmpty(highlight) && !(isMarked && highlight == "#ffff00")) {
                        styles.Add($"background-color:{highlight}");
                    } else if (shading != null) {
                        styles.Add($"background-color:#{shading}");
                    }
                }
                if (styles.Count > 0) {
                    IElement span = CreateOutputElement(htmlDoc, "span");
                    SetOutputAttribute(span, "style", string.Join(";", styles), "RubyRun:style");
                    span.AppendChild(node);
                    node = span;
                }
                string? language = properties.Languages?.Val?.Value;
                if (!string.IsNullOrWhiteSpace(language)) {
                    IElement span = CreateOutputElement(htmlDoc, "span");
                    SetOutputAttribute(span, "lang", language!, "RubyRun:language");
                    span.AppendChild(node);
                    node = span;
                }
                target.AppendChild(node);
            }
        }

        private static INode WrapRubyNode(IDocument htmlDoc, string tagName, INode child) {
            IElement wrapper = CreateOutputElement(htmlDoc, tagName);
            wrapper.AppendChild(child);
            return wrapper;
        }

        private static bool IsEnabled(OnOffType? value) => value != null && value.Val?.Value != false;
    }
}
