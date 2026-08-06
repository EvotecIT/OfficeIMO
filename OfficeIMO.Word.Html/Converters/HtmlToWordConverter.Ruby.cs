using AngleSharp.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private bool TryProcessRubyElement(IElement element, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            if (!TryExtractRubyPairs(element, out var pairs, out bool approximated)) {
                return false;
            }

            currentParagraph ??= AddParagraphInScope(section, cell, headerFooter);
            foreach (RubyPair pair in pairs) {
                currentParagraph._paragraph.AppendChild(CreateRubyContainer(pair, formatting, options));
            }
            if (approximated) {
                AddDiagnostic(
                    options,
                    "HtmlRubyPairingApproximation",
                    "Ruby base and annotation segments could not be paired one-to-one and were combined into one Word ruby run.",
                    "ruby",
                    lossKind: OfficeConversionLossKind.Approximation);
            }
            return true;
        }

        private static Run CreateRubyContainer(RubyPair pair, TextFormatting formatting, HtmlToWordOptions options) =>
            new(new Ruby(
                new RubyProperties(
                    new RubyAlign { Val = RubyAlignValues.Center },
                    new PhoneticGuideTextFontSize { Val = "10" },
                    new PhoneticGuideRaise { Val = 18 },
                    new PhoneticGuideBaseTextSize { Val = "20" },
                    new LanguageId { Val = string.IsNullOrWhiteSpace(formatting.Language) ? "en-US" : formatting.Language! }),
                new RubyContent(CreateRubyRuns(pair.AnnotationElement, pair.RubyText, formatting, options)),
                new RubyBase(CreateRubyRuns(pair.BaseElement, pair.BaseText, formatting, options))));

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

        private static IEnumerable<Run> CreateRubyRuns(
            IElement? element,
            string fallbackText,
            TextFormatting formatting,
            HtmlToWordOptions options) {
            var state = new RubyTextEmissionState();
            if (element == null) {
                Run? fallback = CreateRubyTextRun(fallbackText, formatting, options, state);
                if (fallback != null) yield return fallback;
                yield break;
            }

            TextFormatting rootFormatting = formatting;
            ApplySpanStyles(element, ref rootFormatting);
            List<Run> runs = CreateRubyRuns(element.ChildNodes, rootFormatting, options, state).ToList();
            state.TrimTrailingCollapsibleWhitespace(runs);
            if (runs.Count == 0) {
                Run? fallback = CreateRubyTextRun(fallbackText, rootFormatting, options, state);
                if (fallback != null) runs.Add(fallback);
            }
            foreach (Run run in runs) yield return run;
        }

        private static IEnumerable<Run> CreateRubyRuns(
            IEnumerable<INode> nodes,
            TextFormatting formatting,
            HtmlToWordOptions options,
            RubyTextEmissionState state) {
            foreach (INode node in nodes) {
                if (node is IText textNode) {
                    Run? run = CreateRubyTextRun(textNode.Data, formatting, options, state);
                    if (run != null) yield return run;
                    continue;
                }
                if (node is not IElement childElement || childElement.TagName.Equals("rp", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                TextFormatting childFormatting = formatting;
                ApplyRubySemanticFormatting(childElement.TagName, ref childFormatting);
                ApplySpanStyles(childElement, ref childFormatting);
                foreach (Run childRun in CreateRubyRuns(childElement.ChildNodes, childFormatting, options, state)) {
                    yield return childRun;
                }
            }
        }

        private static Run? CreateRubyTextRun(
            string text,
            TextFormatting formatting,
            HtmlToWordOptions options,
            RubyTextEmissionState state) {
            if (text.Length == 0) return null;

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
            if (prepared.Length == 0) return null;

            Run run = CreateRubyRun(prepared, formatting, options);
            state.HasOutput = true;
            state.PreviousEndedWithCollapsibleWhitespace = collapsible &&
                (prepared[prepared.Length - 1] == ' ' || prepared[prepared.Length - 1] == '\u00A0');
            state.LastCollapsibleText = state.PreviousEndedWithCollapsibleWhitespace
                ? run.GetFirstChild<Text>()
                : null;
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
                case "u": formatting.Underline = true; formatting.UnderlineStyle ??= UnderlineValues.Single; break;
                case "s":
                case "del": formatting.Strike = true; break;
                case "sup": formatting.Superscript = true; formatting.Subscript = false; break;
                case "sub": formatting.Subscript = true; formatting.Superscript = false; break;
                case "small": if (!formatting.FontSize.HasValue) formatting.FontSize = 10; break;
                case "big": if (!formatting.FontSize.HasValue) formatting.FontSize = 18; break;
                case "mark":
                    formatting.Marked = true;
                    formatting.Highlight = HighlightColorValues.Yellow;
                    break;
            }
        }

        private static Run CreateRubyRun(string text, TextFormatting formatting, HtmlToWordOptions options) {
            var run = new Run();
            var properties = CreateRubyRunProperties(formatting, options);
            if (properties.HasChildren || properties.HasAttributes) {
                run.AppendChild(properties);
            }
            run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            return run;
        }

        private static RunProperties CreateRubyRunProperties(TextFormatting formatting, HtmlToWordOptions options) {
            var properties = new RunProperties();
            if (formatting.Bold) properties.AppendChild(new Bold());
            if (formatting.Italic) properties.AppendChild(new Italic());
            if (formatting.Underline) properties.AppendChild(new Underline { Val = GetUnderlineValue(formatting) ?? UnderlineValues.Single });
            if (formatting.Strike) properties.AppendChild(new Strike());
            if (formatting.Superscript) properties.AppendChild(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
            if (formatting.Subscript) properties.AppendChild(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
            if (!string.IsNullOrEmpty(formatting.ColorHex)) properties.AppendChild(new Color { Val = formatting.ColorHex!.TrimStart('#') });
            if (formatting.Highlight.HasValue) properties.AppendChild(new Highlight { Val = formatting.Highlight.Value });
            if (!string.IsNullOrWhiteSpace(formatting.BackgroundColorHex)) {
                if (options.TextBackgroundMode == HtmlTextBackgroundMode.ExactShading) {
                    properties.AppendChild(new Shading {
                        Val = ShadingPatternValues.Clear,
                        Fill = formatting.BackgroundColorHex!.TrimStart('#').ToUpperInvariant()
                    });
                } else if (MapColorToHighlight(formatting.BackgroundColorHex, out _) is HighlightColorValues highlight) {
                    properties.Highlight = new Highlight { Val = highlight };
                }
            }
            if (formatting.Marked) properties.AppendChild(new RunStyle { Val = HtmlSemanticStyleIds.MarkedText });
            if (formatting.FontSize.HasValue) properties.AppendChild(new FontSize { Val = (formatting.FontSize.Value * 2).ToString(CultureInfo.InvariantCulture) });
            if (formatting.Caps == WordCapsStyle.SmallCaps) properties.AppendChild(new SmallCaps());
            if (formatting.Caps == WordCapsStyle.Caps) properties.AppendChild(new Caps());
            if (formatting.LetterSpacing.HasValue) properties.AppendChild(new Spacing { Val = formatting.LetterSpacing.Value });
            if (!string.IsNullOrWhiteSpace(formatting.Language)) properties.AppendChild(new Languages { Val = formatting.Language! });

            var font = ResolveFontFamily(formatting.FontFamily) ?? ResolveFontFamily(options.FontFamily);
            if (!string.IsNullOrEmpty(font)) {
                properties.AppendChild(new RunFonts { Ascii = font, HighAnsi = font, EastAsia = font, ComplexScript = font });
            }

            return properties;
        }

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
            internal Text? LastCollapsibleText { get; set; }

            internal void TrimTrailingCollapsibleWhitespace(List<Run> runs) {
                if (LastCollapsibleText == null) return;

                LastCollapsibleText.Text = LastCollapsibleText.Text.TrimEnd(' ', '\u00A0');
                if (LastCollapsibleText.Text.Length == 0) {
                    Run? emptyRun = runs.FirstOrDefault(run => ReferenceEquals(run.GetFirstChild<Text>(), LastCollapsibleText));
                    if (emptyRun != null) runs.Remove(emptyRun);
                }
                PreviousEndedWithCollapsibleWhitespace = false;
                LastCollapsibleText = null;
            }
        }
    }
}
