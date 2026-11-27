using AngleSharp.Dom;
using System.Globalization;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static bool TryCreateRubyNode(IDocument htmlDoc, WordParagraph run, WordHtmlExportOptions options, out INode node) {
            node = htmlDoc.CreateTextNode(string.Empty);
            OfficeXmlNode ruby = run.RunNode().WChild("ruby");
            if (ruby.IsNone) {
                return false;
            }

            OfficeXmlNode rubyBase = ruby.WChild("rubyBase");
            OfficeXmlNode rubyContent = ruby.WChild("rt");
            var baseText = rubyBase.Text();
            if (string.IsNullOrEmpty(baseText)) {
                return false;
            }

            var rubyText = rubyContent.Text();
            if (string.IsNullOrEmpty(rubyText)) {
                node = htmlDoc.CreateTextNode(baseText);
                return true;
            }

            var rubyElement = CreateOutputElement(htmlDoc, "ruby");
            var baseElement = CreateOutputElement(htmlDoc, "rb");
            AppendRubyRuns(htmlDoc, baseElement, rubyBase, options);
            rubyElement.AppendChild(baseElement);

            var annotationElement = CreateOutputElement(htmlDoc, "rt");
            AppendRubyRuns(htmlDoc, annotationElement, rubyContent, options);
            rubyElement.AppendChild(annotationElement);

            node = rubyElement;
            return true;
        }

        private static void AppendRubyRuns(
            IDocument htmlDoc,
            IElement target,
            OfficeXmlNode source,
            WordHtmlExportOptions options) {
            foreach (OfficeXmlNode run in source.WElements("r")) {
                INode node = htmlDoc.CreateTextNode(run.InnerText);
                OfficeXmlNode properties = run.WChild("rPr");
                if (properties.IsNone) {
                    target.AppendChild(node);
                    continue;
                }
                if (IsEnabled(properties.WChild("b"))) node = WrapRubyNode(htmlDoc, "strong", node);
                if (IsEnabled(properties.WChild("i"))) node = WrapRubyNode(htmlDoc, "em", node);
                if (IsEnabled(properties.WChild("dstrike"))) {
                    var span = CreateOutputElement(htmlDoc, "span");
                    SetOutputAttribute(htmlDoc, span, "style", "text-decoration-line:line-through;text-decoration-style:double", "RubyRunFormatting:double-strike");
                    SetOutputAttribute(htmlDoc, span, "data-officeimo-word-double-strike", "true", "RubyRunFormatting:double-strike-metadata");
                    span.AppendChild(node);
                    node = span;
                } else if (IsEnabled(properties.WChild("strike"))) {
                    node = WrapRubyNode(htmlDoc, "s", node);
                }
                string? underline = properties.WChild("u").WVal();
                if (underline != null && underline != "none") {
                    WordUnderlineStyle wordUnderline = ParseUnderlineToken(underline);
                    if (wordUnderline == WordUnderlineStyle.Single) {
                        node = WrapRubyNode(htmlDoc, "u", node);
                    } else {
                        string cssStyle = MapWordUnderlineToCssStyle(wordUnderline);
                        var span = CreateOutputElement(htmlDoc, "span");
                        SetOutputAttribute(htmlDoc, span, "style", "text-decoration-line:underline;text-decoration-style:" + cssStyle, "RubyRunFormatting:underline");
                        SetOutputAttribute(htmlDoc, span, "data-officeimo-word-underline", wordUnderline.ToString(), "RubyRunFormatting:underline-metadata");
                        span.AppendChild(node);
                        node = span;
                        if (!IsExactCssUnderline(wordUnderline)) {
                            AddWordTextStyleApproximation(options, "word-underline-style-approximated", "Word underline style '" + wordUnderline + "' uses the closest CSS " + cssStyle + " pattern; private round-trip metadata retains the exact Word value.", "word:ruby-run");
                        }
                    }
                }
                string? verticalAlignment = properties.WChild("vertAlign").WVal();
                if (verticalAlignment == "superscript") {
                    node = WrapRubyNode(htmlDoc, "sup", node);
                } else if (verticalAlignment == "subscript") {
                    node = WrapRubyNode(htmlDoc, "sub", node);
                }

                bool isMarked = string.Equals(
                    properties.WChild("rStyle").WVal(),
                    HtmlSemanticStyleIds.MarkedText,
                    StringComparison.Ordinal);
                if (isMarked) node = WrapRubyNode(htmlDoc, "mark", node);

                var styles = new List<string>();
                string? color = NormalizeSixDigitHexColor(properties.WChild("color").WVal());
                if (options.IncludeRunColorStyles && color != null) styles.Add($"color:#{color}");
                if (int.TryParse(properties.WChild("sz").WVal(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int halfPoints)) {
                    styles.Add($"font-size:{(halfPoints / 2D).ToString("0.###", CultureInfo.InvariantCulture)}pt");
                }
                string? font = properties.WChild("rFonts").WAttr("ascii") ?? properties.WChild("rFonts").WAttr("hAnsi");
                if (options.IncludeFontStyles && !string.IsNullOrWhiteSpace(font)) styles.Add($"font-family:{FontFamilyCss(font!)}");
                if (IsEnabled(properties.WChild("smallCaps"))) styles.Add("font-variant:small-caps");
                if (IsEnabled(properties.WChild("caps"))) styles.Add("text-transform:uppercase");
                if (properties.WChild("spacing").WInt32("val") is int twentiethPoints) {
                    styles.Add($"letter-spacing:{(twentiethPoints / 20D).ToString("0.###", CultureInfo.InvariantCulture)}pt");
                }
                if (options.IncludeRunHighlightStyles) {
                    string? shading = NormalizeSixDigitHexColor(properties.WChild("shd").WAttr("fill"));
                    string? highlight = GetHighlightCss(properties.WChild("highlight").WVal());
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
                string? language = properties.WChild("lang").WVal();
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

        private static bool IsEnabled(OfficeXmlNode value) => value.Exists && WordXml.ParseOnOff(value.WVal()) != false;

        private static bool IsExplicitlyDisabled(OfficeXmlNode value) => value.Exists && WordXml.ParseOnOff(value.WVal()) == false;

        private static WordUnderlineStyle ParseUnderlineToken(string token) =>
            WordUnderlineStyleTokens.TryParse(token, out WordUnderlineStyle style)
                ? style
                : throw new ArgumentOutOfRangeException(nameof(token), token, "Unsupported Word underline value.");

        private static WordVerticalTextPosition ParseVerticalPositionToken(string token) =>
            WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordVerticalTextPosition position)
                ? position
                : throw new ArgumentOutOfRangeException(nameof(token), token, "Unsupported Word vertical position value.");
    }
}
