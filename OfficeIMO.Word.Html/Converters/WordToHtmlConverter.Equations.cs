using AngleSharp.Dom;
using AngleSharp.Html;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private static IElement? CreateEquationNode(
            IHtmlDocument htmlDocument,
            IElement context,
            WordEquation equation,
            WordToHtmlOptions options) {
            string label = equation.Text;
            if (equation.Representation == WordEquationRepresentation.EquationField) {
                // InspectExport charged the cached field result as ordinary Word text. The
                // MathML projection replaces that source text, so release it before charging
                // the generated mtext and accessibility label that actually reach the output.
                ReleaseOutputCharacters(
                    htmlDocument,
                    GetHtmlEncodedLength(label, attributeValue: false));
            }
            long remaining = GetRemainingOutputCharacters(htmlDocument);
            string mathMl;
            try {
                mathMl = equation.ToMathMl(GetMathMlParserInputLimit(remaining));
            } catch (WordMathMlOutputLimitExceededException) {
                ThrowExportLimitExceeded(
                    options,
                    "WordHtmlOutputLimitExceeded",
                    "Generated MathML exceeds the configured HTML output-character limit before DOM construction.",
                    "EquationMathMl",
                    SaturatingAdd(options.MaxOutputCharacters, 1),
                    options.MaxOutputCharacters);
                return null;
            }
            IElement? mathNode = new HtmlParser()
                .ParseFragment(mathMl, context)
                .OfType<IElement>()
                .FirstOrDefault();
            if (mathNode == null) return null;
            using var countingWriter = new CountingHtmlWriter();
            mathNode.ToHtml(countingWriter, HtmlMarkupFormatter.Instance);
            ReserveOutputCharacters(
                htmlDocument,
                countingWriter.CharacterCount,
                "Generated MathML exceeds the configured HTML output-character limit before DOM construction.",
                "EquationMathMl");
            ReserveOutputCharacters(
                htmlDocument,
                " aria-label=\"\"".Length + GetHtmlEncodedLength(label, attributeValue: true),
                "Generated equation accessibility metadata exceeds the configured HTML output-character limit before DOM construction.",
                "EquationAriaLabel");
            // The complete serialized aria-label attribute is reserved before DOM assignment.
            mathNode.SetAttribute("aria-label", label);
            return mathNode;
        }

        private static long GetMathMlParserInputLimit(long remainingOutputCharacters) =>
            remainingOutputCharacters > long.MaxValue / 6L
                ? long.MaxValue
                : remainingOutputCharacters * 6L;

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
                var strong = CreateOutputElement(htmlDocument, "strong");
                strong.AppendChild(node);
                node = strong;
            }
            if (run.Italic) {
                var emphasis = CreateOutputElement(htmlDocument, "em");
                emphasis.AppendChild(node);
                node = emphasis;
            }
            if ((run.Strike || run.DoubleStrike) && !isHtmlDeletedText) {
                var strike = CreateOutputElement(htmlDocument, "s");
                strike.AppendChild(node);
                node = strike;
            }
            if (run.Underline != null && !isHtmlInsertedText) {
                var underline = CreateOutputElement(htmlDocument, "u");
                underline.AppendChild(node);
                node = underline;
            }
            if (run.VerticalTextAlignment == VerticalPositionValues.Superscript) {
                var superscript = CreateOutputElement(htmlDocument, "sup");
                superscript.AppendChild(node);
                node = superscript;
            } else if (run.VerticalTextAlignment == VerticalPositionValues.Subscript) {
                var subscript = CreateOutputElement(htmlDocument, "sub");
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
                options.IncludeRunHighlightStyles,
                out bool handledHtmlStyle);
            if (options.IncludeFontStyles) {
                string? font = run.FontFamily ?? options.FontFamily;
                if (!string.IsNullOrEmpty(font)) {
                    var span = CreateOutputElement(htmlDocument, "span");
                    SetOutputAttribute(
                        htmlDocument,
                        span,
                        "style",
                        $"font-family:{QuoteCssString(font!)}",
                        "EquationRunFontStyle");
                    span.AppendChild(node);
                    node = span;
                }
            }
            if (run.FontSize != null) {
                var span = CreateOutputElement(htmlDocument, "span");
                SetOutputAttribute(span, "style", $"font-size:{run.FontSize.Value}pt", "EquationRunFormatting:font-size");
                span.AppendChild(node);
                node = span;
            }
            if (run.CapsStyle == CapsStyle.SmallCaps || run.CapsStyle == CapsStyle.Caps) {
                var span = CreateOutputElement(htmlDocument, "span");
                SetOutputAttribute(
                    span,
                    "style",
                    run.CapsStyle == CapsStyle.SmallCaps
                        ? "font-variant:small-caps"
                        : "text-transform:uppercase",
                    "EquationRunFormatting:caps");
                span.AppendChild(node);
                node = span;
            }
            if (options.IncludeRunColorStyles || options.IncludeRunHighlightStyles) {
                var styles = new List<string>();
                if (options.IncludeRunColorStyles &&
                    !string.IsNullOrEmpty(run.ColorHex) &&
                    !string.Equals(run.ColorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                    string? normalized = NormalizeSixDigitHexColor(run.ColorHex);
                    if (normalized != null) styles.Add($"color:#{normalized}");
                }
                if (options.IncludeRunHighlightStyles && !isHtmlMarkedText) {
                    string? normalizedRunBackground = NormalizeSixDigitHexColor(
                        WordDocumentImageRenderer.ResolveRunShadingFillColorHex(run));
                    string? highlight = GetHighlightCss(
                        WordDocumentImageRenderer.ResolveRunHighlight(run));
                    if (!string.IsNullOrEmpty(highlight) &&
                        (!isHtmlMarkedText || normalizedRunBackground != null)) {
                        styles.Add($"background-color:{highlight}");
                    } else if (normalizedRunBackground != null) {
                        styles.Add($"background-color:#{normalizedRunBackground}");
                    }
                }
                if (styles.Count > 0) {
                    var span = CreateOutputElement(htmlDocument, "span");
                    SetOutputAttribute(span, "style", string.Join(";", styles), "EquationRunFormatting:color-highlight");
                    span.AppendChild(node);
                    node = span;
                }
            }
            if (options.IncludeRunClasses && !string.IsNullOrEmpty(run.CharacterStyleId) && !handledHtmlStyle) {
                var span = CreateOutputElement(htmlDocument, "span");
                SetOutputAttribute(span, "class", GetSafeStyleClassName(run.CharacterStyleId), "EquationRunFormatting:class");
                span.AppendChild(node);
                node = span;
                runStyles.Add(run.CharacterStyleId!);
            }
            string? language = NormalizeRunLanguage(run.Language, documentLanguage);
            if (!string.IsNullOrEmpty(language)) {
                var span = CreateOutputElement(htmlDocument, "span");
                SetOutputAttribute(span, "lang", language!, "EquationRunFormatting:language");
                span.AppendChild(node);
                node = span;
            }
            return node;
        }

        private INode ApplyHtmlSemanticCharacterStyle(
            IHtmlDocument htmlDocument,
            WordParagraph run,
            string text,
            INode node,
            bool includeRunHighlightStyles,
            out bool handled) {
            handled = true;
            IElement semanticNode;
            if (string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.DeletedText, StringComparison.OrdinalIgnoreCase)) {
                semanticNode = CreateOutputElement(htmlDocument, "del");
            } else if (string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.InsertedText, StringComparison.OrdinalIgnoreCase)) {
                semanticNode = CreateOutputElement(htmlDocument, "ins");
            } else if (string.Equals(run.CharacterStyleId, HtmlSemanticStyleIds.MarkedText, StringComparison.OrdinalIgnoreCase)) {
                semanticNode = CreateOutputElement(htmlDocument, "mark");
                string? normalizedRunBackground = includeRunHighlightStyles
                    ? NormalizeSixDigitHexColor(WordDocumentImageRenderer.ResolveRunShadingFillColorHex(run))
                    : null;
                string? highlightCss = includeRunHighlightStyles
                    ? GetHighlightCss(WordDocumentImageRenderer.ResolveRunHighlight(run))
                    : null;
                bool isDefaultMarkHighlight = string.Equals(
                    highlightCss,
                    "#ffff00",
                    StringComparison.OrdinalIgnoreCase);
                if (!string.IsNullOrEmpty(highlightCss) &&
                    (normalizedRunBackground != null || !isDefaultMarkHighlight)) {
                    SetOutputAttribute(semanticNode, "style", $"background-color:{highlightCss}", "EquationSemanticFormatting:highlight");
                } else if (normalizedRunBackground != null) {
                    SetOutputAttribute(semanticNode, "style", $"background-color:#{normalizedRunBackground}", "EquationSemanticFormatting:highlight");
                } else if (run.Highlight == HighlightColorValues.None) {
                    SetOutputAttribute(semanticNode, "style", "background-color:transparent", "EquationSemanticFormatting:highlight");
                }
            } else if (string.Equals(run.CharacterStyleId, "HtmlCite", StringComparison.OrdinalIgnoreCase)) {
                semanticNode = CreateOutputElement(htmlDocument, "cite");
            } else if (string.Equals(run.CharacterStyleId, "HtmlDfn", StringComparison.OrdinalIgnoreCase)) {
                semanticNode = CreateOutputElement(htmlDocument, "dfn");
            } else if (string.Equals(run.CharacterStyleId, "HtmlTime", StringComparison.OrdinalIgnoreCase)) {
                semanticNode = CreateOutputElement(htmlDocument, "time");
                bool hasImportedDateTime = HtmlSemanticMetadata.TryGetTimeDateTime(run, out string dateTime);
                if (!hasImportedDateTime) {
                    dateTime = text;
                    if (DateTime.TryParse(text, out DateTime parsed)) {
                        dateTime = parsed.ToString("o");
                    }
                }
                SetOutputAttribute(semanticNode, "datetime", dateTime, "EquationSemanticFormatting:datetime");
            } else if (string.Equals(run.CharacterStyleId, "HtmlCode", StringComparison.OrdinalIgnoreCase)) {
                semanticNode = CreateOutputElement(htmlDocument, "code");
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

            IElement anchor = CreateOutputElement(htmlDocument, "a");
            SetOutputAttribute(htmlDocument, anchor, "href", href!, "EquationHyperlink:href");
            if (!string.IsNullOrEmpty(hyperlink.Tooltip)) {
                SetOutputAttribute(htmlDocument, anchor, "title", hyperlink.Tooltip!, "EquationHyperlink:title");
            }
            string? targetFrame = hyperlink._hyperlink.TargetFrame?.Value;
            if (!string.IsNullOrEmpty(targetFrame)) {
                SetOutputAttribute(htmlDocument, anchor, "target", targetFrame!, "EquationHyperlink:target");
            }
            return anchor;
        }
    }
}
