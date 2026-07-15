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
            ISet<string> runStyles) {
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
            if (run.Strike || run.DoubleStrike) {
                var strike = htmlDocument.CreateElement("s");
                strike.AppendChild(node);
                node = strike;
            }
            if (run.Underline != null) {
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
            if (run.IsHyperLink && run.Hyperlink != null) {
                string? href = run.Hyperlink.Uri?.ToString();
                if (string.IsNullOrEmpty(href) && !string.IsNullOrEmpty(run.Hyperlink.Anchor)) {
                    href = "#" + run.Hyperlink.Anchor;
                }
                if (!string.IsNullOrEmpty(href)) {
                    var anchor = htmlDocument.CreateElement("a");
                    anchor.SetAttribute("href", href);
                    anchor.AppendChild(node);
                    node = anchor;
                }
            }
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
                if (options.IncludeRunHighlightStyles) {
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
            if (options.IncludeRunClasses && !string.IsNullOrEmpty(run.CharacterStyleId)) {
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
    }
}
