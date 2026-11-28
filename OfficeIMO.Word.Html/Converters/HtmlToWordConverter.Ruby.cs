using AngleSharp.Dom;
using OfficeIMO.Html;
using OfficeIMO.OpenXml.Internal;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private bool TryProcessRubyElement(IElement element, WordSection section, HtmlWordExportOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            if (!TryExtractRubyPairs(element, out var pairs, out bool approximated)) {
                return false;
            }

            currentParagraph ??= AddParagraphInScope(section, cell, headerFooter);
            OfficeXmlNode paragraphNode = currentParagraph.Node();
            foreach (RubyPair pair in pairs) {
                paragraphNode.Append(CreateRubyContainer(pair, formatting, options));
            }
            if (approximated) {
                AddDiagnostic(
                    options,
                    "html-ruby-pairing-approximation",
                    "Ruby base and annotation segments could not be paired one-to-one and were combined into one Word ruby run.",
                    "ruby",
                    lossKind: OfficeConversionLossKind.Approximation);
            }
            return true;
        }

        private static OfficeXmlNode CreateRubyContainer(RubyPair pair, TextFormatting formatting, HtmlWordExportOptions options) {
            OfficeXmlNode container = WordConverterNodes.CreateDetachedRun();
            OfficeXmlNode ruby = container.Append(container.CreateElement(WordXml.W, "ruby"));
            OfficeXmlNode properties = ruby.Append(ruby.CreateElement(WordXml.W, "rubyPr"));
            AppendValueElement(properties, "rubyAlign", "center");
            AppendValueElement(properties, "hps", "10");
            AppendValueElement(properties, "hpsRaise", "18");
            AppendValueElement(properties, "hpsBaseText", "20");
            AppendValueElement(properties, "lid", string.IsNullOrWhiteSpace(formatting.Language) ? "en-US" : formatting.Language!);
            OfficeXmlNode content = ruby.Append(ruby.CreateElement(WordXml.W, "rt"));
            foreach (OfficeXmlNode run in CreateRubyRuns(pair.AnnotationElement, pair.RubyText, formatting, options)) content.Append(run);
            OfficeXmlNode rubyBase = ruby.Append(ruby.CreateElement(WordXml.W, "rubyBase"));
            foreach (OfficeXmlNode run in CreateRubyRuns(pair.BaseElement, pair.BaseText, formatting, options)) rubyBase.Append(run);
            return container;
        }

        /// <summary>Appends a <c>w:</c> element with a <c>w:val</c> to <paramref name="parent"/> and returns it.</summary>
        private static OfficeXmlNode AppendValueElement(OfficeXmlNode parent, string localName, string value) {
            OfficeXmlNode element = parent.Append(parent.CreateElement(WordXml.W, localName));
            element.SetAttribute(WordXml.W, "val", value);
            return element;
        }

        private static bool TryExtractRubyPairs(IElement element, out List<RubyPair> pairs, out bool approximated) {
            pairs = new List<RubyPair>();
            approximated = false;
            var explicitBases = element.Children
                .Where(child => child.TagName.Equals("rb", StringComparison.OrdinalIgnoreCase))
                .Where(child => !string.IsNullOrWhiteSpace(child.TextContent))
                .ToList();
            var explicitAnnotations = element.Children
                .Where(child => child.TagName.Equals("rt", StringComparison.OrdinalIgnoreCase))
                .Where(child => !string.IsNullOrWhiteSpace(child.TextContent))
                .ToList();
            if (explicitBases.Count > 0 && explicitBases.Count == explicitAnnotations.Count) {
                for (int index = 0; index < explicitBases.Count; index++) {
                    pairs.Add(new RubyPair(
                        explicitBases[index],
                        explicitAnnotations[index],
                        NormalizeRubyText(explicitBases[index].TextContent),
                        NormalizeRubyText(explicitAnnotations[index].TextContent)));
                }
                return true;
            }

            var baseBuilder = new StringBuilder();
            var rubyBuilder = new StringBuilder();

            foreach (var child in element.ChildNodes) {
                if (child is IElement childElement) {
                    var tagName = childElement.TagName.ToLowerInvariant();
                    if (tagName == "rt") {
                        rubyBuilder.Append(childElement.TextContent);
                    } else if (tagName == "rp") {
                        continue;
                    } else if (tagName == "rb") {
                        baseBuilder.Append(childElement.TextContent);
                    } else {
                        baseBuilder.Append(childElement.TextContent);
                    }
                } else {
                    baseBuilder.Append(child.TextContent);
                }
            }

            string baseText = NormalizeRubyText(baseBuilder.ToString());
            string rubyText = NormalizeRubyText(rubyBuilder.ToString());
            if (string.IsNullOrWhiteSpace(baseText) || string.IsNullOrWhiteSpace(rubyText)) {
                return false;
            }

            pairs.Add(new RubyPair(null, null, baseText, rubyText));
            approximated = explicitBases.Count > 1 || explicitAnnotations.Count > 1 ||
                           explicitBases.Count != explicitAnnotations.Count;
            return true;
        }

        private static string NormalizeRubyText(string text) =>
            string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

        private static IEnumerable<OfficeXmlNode> CreateRubyRuns(
            IElement? element,
            string fallbackText,
            TextFormatting formatting,
            HtmlWordExportOptions options) {
            var state = new RubyTextEmissionState();
            if (element == null) {
                OfficeXmlNode fallback = CreateRubyTextRun(fallbackText, formatting, options, state);
                if (fallback.Exists) yield return fallback;
                yield break;
            }

            TextFormatting rootFormatting = formatting;
            ApplySpanStyles(element, ref rootFormatting);
            List<OfficeXmlNode> runs = CreateRubyRuns(element.ChildNodes, rootFormatting, options, state).ToList();
            state.TrimTrailingCollapsibleWhitespace(runs);
            if (runs.Count == 0) {
                OfficeXmlNode fallback = CreateRubyTextRun(fallbackText, rootFormatting, options, state);
                if (fallback.Exists) runs.Add(fallback);
            }
            foreach (OfficeXmlNode run in runs) yield return run;
        }

        private static IEnumerable<OfficeXmlNode> CreateRubyRuns(
            IEnumerable<INode> nodes,
            TextFormatting formatting,
            HtmlWordExportOptions options,
            RubyTextEmissionState state) {
            foreach (INode node in nodes) {
                if (node is IText textNode) {
                    OfficeXmlNode run = CreateRubyTextRun(textNode.Data, formatting, options, state);
                    if (run.Exists) yield return run;
                    continue;
                }
                if (node is not IElement childElement || childElement.TagName.Equals("rp", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                TextFormatting childFormatting = formatting;
                ApplyRubySemanticFormatting(childElement.TagName, ref childFormatting);
                ApplySpanStyles(childElement, ref childFormatting);
                foreach (OfficeXmlNode childRun in CreateRubyRuns(childElement.ChildNodes, childFormatting, options, state)) {
                    yield return childRun;
                }
            }
        }

        private static OfficeXmlNode CreateRubyTextRun(
            string text,
            TextFormatting formatting,
            HtmlWordExportOptions options,
            RubyTextEmissionState state) {
            if (text.Length == 0) return default;

            WhiteSpaceMode mode = formatting.WhiteSpace ?? WhiteSpaceMode.Normal;
            bool collapsible = mode is WhiteSpaceMode.Normal or WhiteSpaceMode.NoWrap;
            string prepared = ApplyWhiteSpace(text, mode);
            if (collapsible) {
                char space = mode == WhiteSpaceMode.NoWrap ? '\u00A0' : ' ';
                if (!state.HasOutput || state.PreviousEndedWithCollapsibleWhitespace) {
                    prepared = prepared.TrimStart(space);
                }
            }
            prepared = ApplyRubyTextTransform(prepared, formatting.Transform, state);
            if (prepared.Length == 0) return default;

            OfficeXmlNode run = CreateRubyRun(prepared, formatting, options);
            state.HasOutput = true;
            state.PreviousEndedWithCollapsibleWhitespace = collapsible &&
                (prepared[prepared.Length - 1] == ' ' || prepared[prepared.Length - 1] == '\u00A0');
            state.LastCollapsibleText = state.PreviousEndedWithCollapsibleWhitespace
                ? run.WChild("t")
                : default;
            return run;
        }

        private static string ApplyRubyTextTransform(
            string text,
            TextTransform transform,
            RubyTextEmissionState state) {
            string transformed;
            if (transform == TextTransform.Capitalize && state.PreviousEndedInsideWord) {
                string continuedWord = "a" + text;
                transformed = CultureInfo.CurrentCulture.TextInfo
                    .ToTitleCase(continuedWord.ToLowerInvariant())
                    .Substring(1);
            } else {
                transformed = ApplyTextTransform(text, transform);
            }

            state.PreviousEndedInsideWord = EndsInsideRubyWord(text);
            return transformed;
        }

        private static bool EndsInsideRubyWord(string text) {
            for (int index = text.Length - 1; index >= 0;) {
                if (char.IsLowSurrogate(text[index]) && index > 0 && char.IsHighSurrogate(text[index - 1])) {
                    index--;
                }
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(text, index);
                if (category is UnicodeCategory.NonSpacingMark or
                    UnicodeCategory.SpacingCombiningMark or
                    UnicodeCategory.EnclosingMark) {
                    index--;
                    continue;
                }
                return category is UnicodeCategory.UppercaseLetter or
                    UnicodeCategory.LowercaseLetter or
                    UnicodeCategory.TitlecaseLetter or
                    UnicodeCategory.ModifierLetter or
                    UnicodeCategory.OtherLetter or
                    UnicodeCategory.DecimalDigitNumber or
                    UnicodeCategory.LetterNumber or
                    UnicodeCategory.OtherNumber;
            }
            return false;
        }

        private static void ApplyRubySemanticFormatting(string tagName, ref TextFormatting formatting) {
            switch (tagName.ToLowerInvariant()) {
                case "strong":
                case "b": formatting.Bold = true; break;
                case "em":
                case "i": formatting.Italic = true; break;
                case "u": formatting.Underline = true; formatting.UnderlineStyle ??= WordUnderlineStyle.Single; break;
                case "s":
                case "del": formatting.Strike = true; break;
                case "sup": formatting.Superscript = true; formatting.Subscript = false; break;
                case "sub": formatting.Subscript = true; formatting.Superscript = false; break;
                case "small": if (!formatting.FontSize.HasValue) formatting.FontSize = 10; break;
                case "big": if (!formatting.FontSize.HasValue) formatting.FontSize = 18; break;
                case "mark":
                    formatting.Marked = true;
                    formatting.Highlight = WordHighlightColor.Yellow;
                    break;
            }
        }

        private static OfficeXmlNode CreateRubyRun(string text, TextFormatting formatting, HtmlWordExportOptions options) {
            OfficeXmlNode run = WordConverterNodes.CreateDetachedRun();
            OfficeXmlNode properties = CreateRubyRunProperties(run, formatting, options);
            if (properties.HasChildren() || properties.HasAttributes) {
                run.Append(properties);
            }
            run.Append(WordConverterNodes.CreateDetachedPreservedText(text));
            return run;
        }

        // Properties the SDK set through typed members go in schema order; the rest are appended in this order.
        private static OfficeXmlNode CreateRubyRunProperties(OfficeXmlNode run, TextFormatting formatting, HtmlWordExportOptions options) {
            OfficeXmlNode properties = run.CreateElement(WordXml.W, "rPr");
            if (formatting.Bold.HasValue) WordXmlEdits.ReplaceChild(properties, "b", WRPrView.Order, OnOffToken(formatting.Bold.Value));
            if (formatting.Italic.HasValue) WordXmlEdits.ReplaceChild(properties, "i", WRPrView.Order, OnOffToken(formatting.Italic.Value));
            if (formatting.Underline.HasValue) WordXmlEdits.ReplaceChild(properties, "u", WRPrView.Order, UnderlineStyleToken(formatting.Underline.Value
                ? GetUnderlineValue(formatting) ?? WordUnderlineStyle.Single : WordUnderlineStyle.None));
            if (formatting.Strike.HasValue) WordXmlEdits.ReplaceChild(properties, "strike", WRPrView.Order, OnOffToken(formatting.Strike.Value));
            if (formatting.DoubleStrike.HasValue) WordXmlEdits.ReplaceChild(properties, "dstrike", WRPrView.Order, OnOffToken(formatting.DoubleStrike.Value));
            if (formatting.Superscript) AppendValueElement(properties, "vertAlign", "superscript");
            if (formatting.Subscript) AppendValueElement(properties, "vertAlign", "subscript");
            if (!string.IsNullOrEmpty(formatting.ColorHex)) AppendValueElement(properties, "color", formatting.ColorHex!.TrimStart('#'));
            if (formatting.Highlight.HasValue) AppendValueElement(properties, "highlight", formatting.Highlight.Value.ToOpenXmlToken());
            if (!string.IsNullOrWhiteSpace(formatting.BackgroundColorHex)) {
                if (options.TextBackgroundMode == WordHtmlTextBackgroundMode.ExactShading) {
                    OfficeXmlNode shading = AppendValueElement(properties, "shd", "clear");
                    shading.SetAttribute(WordXml.W, "fill", formatting.BackgroundColorHex!.TrimStart('#').ToUpperInvariant());
                } else if (MapColorToHighlight(formatting.BackgroundColorHex, out _) is WordHighlightColor highlight) {
                    WordXmlEdits.ReplaceChild(properties, "highlight", WRPrView.Order, highlight.ToOpenXmlToken());
                }
            }
            if (formatting.Marked) AppendValueElement(properties, "rStyle", HtmlSemanticStyleIds.MarkedText);
            if (formatting.FontSize.HasValue) AppendValueElement(properties, "sz", ((int)Math.Round(formatting.FontSize.Value * 2d, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture));
            if (formatting.Caps == WordCapsStyle.SmallCaps) properties.Append(properties.CreateElement(WordXml.W, "smallCaps"));
            if (formatting.Caps == WordCapsStyle.Caps) properties.Append(properties.CreateElement(WordXml.W, "caps"));
            if (formatting.LetterSpacing.HasValue) AppendValueElement(properties, "spacing", formatting.LetterSpacing.Value.ToString(CultureInfo.InvariantCulture));
            if (!string.IsNullOrWhiteSpace(formatting.Language)) AppendValueElement(properties, "lang", formatting.Language!);

            var font = ResolveFontFamily(formatting.FontFamily) ?? ResolveFontFamily(GetRunFontFamilyOption(options));
            if (!string.IsNullOrEmpty(font)) {
                OfficeXmlNode fonts = properties.Append(properties.CreateElement(WordXml.W, "rFonts"));
                fonts.SetAttribute(WordXml.W, "ascii", font);
                fonts.SetAttribute(WordXml.W, "hAnsi", font);
                fonts.SetAttribute(WordXml.W, "eastAsia", font);
                fonts.SetAttribute(WordXml.W, "cs", font);
            }

            return properties;
        }

        private static string OnOffToken(bool value) => value ? "true" : "false";

        private readonly struct RubyPair {
            internal RubyPair(IElement? baseElement, IElement? annotationElement, string baseText, string rubyText) {
                BaseElement = baseElement;
                AnnotationElement = annotationElement;
                BaseText = baseText;
                RubyText = rubyText;
            }

            internal IElement? BaseElement { get; }
            internal IElement? AnnotationElement { get; }
            internal string BaseText { get; }
            internal string RubyText { get; }
        }

        private sealed class RubyTextEmissionState {
            internal bool HasOutput { get; set; }
            internal bool PreviousEndedWithCollapsibleWhitespace { get; set; }
            internal bool PreviousEndedInsideWord { get; set; }
            internal OfficeXmlNode LastCollapsibleText { get; set; }

            internal void TrimTrailingCollapsibleWhitespace(List<OfficeXmlNode> runs) {
                OfficeXmlNode text = LastCollapsibleText;
                if (text.IsNone) return;

                text.SetInnerText(text.InnerText.TrimEnd(' ', '\u00A0'));
                if (text.InnerText.Length == 0) {
                    int emptyRun = runs.FindIndex(run => run.WChild("t") == text);
                    if (emptyRun >= 0) runs.RemoveAt(emptyRun);
                }
                PreviousEndedWithCollapsibleWhitespace = false;
                LastCollapsibleText = default;
            }
        }
    }
}
