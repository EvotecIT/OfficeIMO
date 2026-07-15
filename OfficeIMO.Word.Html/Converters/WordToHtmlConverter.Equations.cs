using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private INode CreateEquationAdjacentTextNode(
            IHtmlDocument htmlDocument,
            WordParagraph run,
            string text,
            WordToHtmlOptions options,
            string? documentLanguage,
            ISet<string> runStyles,
            bool includeHyperlink = true) {
            bool isHtmlDeletedText = string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.DeletedText, StringComparison.OrdinalIgnoreCase);
            bool isHtmlInsertedText = string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.InsertedText, StringComparison.OrdinalIgnoreCase);
            bool isHtmlMarkedText = string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.MarkedText, StringComparison.OrdinalIgnoreCase);
            INode node = htmlDocument.CreateTextNode(text);
            if (run.Bold) {
                var strong = htmlDocument.CreateElement("strong");
                strong.AppendChild(node);
                node = strong;
            }
            if (run.Italic) {
                var emphasis = htmlDocument.CreateElement("em");
                emphasis.AppendChild(node);
                node = emphasis;
            }
            if ((run.Strike || run.DoubleStrike) && !isHtmlDeletedText) {
                var strike = htmlDocument.CreateElement("s");
                strike.AppendChild(node);
                node = strike;
            }
            if (run.Underline != null && !isHtmlInsertedText) {
                var underline = htmlDocument.CreateElement("u");
                underline.AppendChild(node);
                node = underline;
            }
            if (run.VerticalTextAlignment == VerticalPositionValues.Superscript) {
                var superscript = htmlDocument.CreateElement("sup");
                superscript.AppendChild(node);
                node = superscript;
            } else if (run.VerticalTextAlignment == VerticalPositionValues.Subscript) {
                var subscript = htmlDocument.CreateElement("sub");
                subscript.AppendChild(node);
                node = subscript;
            }
            if (includeHyperlink && run.IsHyperLink && run.Hyperlink != null) {
                IElement? anchor = CreateEquationHyperlinkNode(htmlDocument, run.Hyperlink);
                if (anchor != null) {
                    anchor.AppendChild(node);
                    node = anchor;
                }
            }
            node = ApplyHtmlSemanticCharacterStyle(
                htmlDocument,
                run,
                text,
                node,
                out bool handledHtmlStyle);
            if (options.IncludeFontStyles) {
                string? font = run.FontFamily ?? options.FontFamily;
                if (!string.IsNullOrEmpty(font)) {
                    var span = htmlDocument.CreateElement("span");
                    string value = font!.Contains(' ') ? $"\"{font}\"" : font!;
                    span.SetAttribute("style", $"font-family:{value}");
                    span.AppendChild(node);
                    node = span;
                }
            }
            if (run.FontSize != null) {
                var span = htmlDocument.CreateElement("span");
                span.SetAttribute("style", $"font-size:{run.FontSize.Value}pt");
                span.AppendChild(node);
                node = span;
            }
            if (run.CapsStyle == CapsStyle.SmallCaps || run.CapsStyle == CapsStyle.Caps) {
                var span = htmlDocument.CreateElement("span");
                span.SetAttribute("style", run.CapsStyle == CapsStyle.SmallCaps
                    ? "font-variant:small-caps"
                    : "text-transform:uppercase");
                span.AppendChild(node);
                node = span;
            }
            if (options.IncludeRunColorStyles || options.IncludeRunHighlightStyles) {
                var styles = new List<string>();
                if (options.IncludeRunColorStyles &&
                    !string.IsNullOrEmpty(run.ColorHex) &&
                    !string.Equals(run.ColorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                    styles.Add($"color:#{run.ColorHex.Trim().TrimStart('#').ToLowerInvariant()}");
                }
                if (options.IncludeRunHighlightStyles && !isHtmlMarkedText) {
                    string? highlight = GetHighlightCss(run.Highlight);
                    if (!string.IsNullOrEmpty(highlight)) styles.Add($"background-color:{highlight}");
                }
                if (styles.Count > 0) {
                    var span = htmlDocument.CreateElement("span");
                    span.SetAttribute("style", string.Join(";", styles));
                    span.AppendChild(node);
                    node = span;
                }
            }
            if (options.IncludeRunClasses && !string.IsNullOrEmpty(run.CharacterStyleId) && !handledHtmlStyle) {
                var span = htmlDocument.CreateElement("span");
                span.SetAttribute("class", run.CharacterStyleId);
                span.AppendChild(node);
                node = span;
                runStyles.Add(run.CharacterStyleId!);
            }
            string? language = NormalizeRunLanguage(run.Language, documentLanguage);
            if (!string.IsNullOrEmpty(language)) {
                var span = htmlDocument.CreateElement("span");
                span.SetAttribute("lang", language);
                span.AppendChild(node);
                node = span;
            }
            return node;
        }

        private static INode ApplyHtmlSemanticCharacterStyle(
            IHtmlDocument htmlDocument,
            WordParagraph run,
            string text,
            INode node,
            out bool handled) {
            handled = true;
            IElement semanticNode;
            if (string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.DeletedText, StringComparison.OrdinalIgnoreCase)) {
                semanticNode = htmlDocument.CreateElement("del");
            } else if (string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.InsertedText, StringComparison.OrdinalIgnoreCase)) {
                semanticNode = htmlDocument.CreateElement("ins");
            } else if (string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.MarkedText, StringComparison.OrdinalIgnoreCase)) {
                semanticNode = htmlDocument.CreateElement("mark");
            } else if (string.Equals(run.CharacterStyleId, "HtmlCite", StringComparison.OrdinalIgnoreCase)) {
                semanticNode = htmlDocument.CreateElement("cite");
            } else if (string.Equals(run.CharacterStyleId, "HtmlDfn", StringComparison.OrdinalIgnoreCase)) {
                semanticNode = htmlDocument.CreateElement("dfn");
            } else if (string.Equals(run.CharacterStyleId, "HtmlTime", StringComparison.OrdinalIgnoreCase)) {
                semanticNode = htmlDocument.CreateElement("time");
                bool hasImportedDateTime = HtmlSemanticMetadata.TryGetTimeDateTime(run, out string dateTime);
                if (!hasImportedDateTime) {
                    dateTime = text;
                    if (DateTime.TryParse(text, out DateTime parsed)) {
                        dateTime = parsed.ToString("o");
                    }
                }
                semanticNode.SetAttribute("datetime", dateTime);
            } else if (string.Equals(run.CharacterStyleId, "HtmlCode", StringComparison.OrdinalIgnoreCase)) {
                semanticNode = htmlDocument.CreateElement("code");
            } else {
                handled = false;
                return node;
            }

            semanticNode.AppendChild(node);
            return semanticNode;
        }

        private static IElement? CreateEquationHyperlinkNode(IHtmlDocument htmlDocument, WordHyperLink hyperlink) {
            string? href = hyperlink.Uri?.ToString();
            if (string.IsNullOrEmpty(href) && !string.IsNullOrEmpty(hyperlink.Anchor)) {
                href = "#" + hyperlink.Anchor;
            }
            if (string.IsNullOrEmpty(href)) {
                return null;
            }

            IElement anchor = htmlDocument.CreateElement("a");
            anchor.SetAttribute("href", href);
            if (!string.IsNullOrEmpty(hyperlink.Tooltip)) {
                anchor.SetAttribute("title", hyperlink.Tooltip);
            }
            string? targetFrame = hyperlink._hyperlink.TargetFrame?.Value;
            if (!string.IsNullOrEmpty(targetFrame)) {
                anchor.SetAttribute("target", targetFrame);
            }
            return anchor;
        }
    }
}
