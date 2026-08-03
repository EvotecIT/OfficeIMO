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
            foreach (var pair in pairs) {
                currentParagraph._paragraph.AppendChild(CreateRubyContainer(pair.BaseText, pair.RubyText, formatting, options));
            }
            if (approximated) {
                AddDiagnostic(
                    options,
                    "HtmlRubyPairingApproximation",
                    "Ruby base and annotation segments could not be paired one-to-one and were combined into one Word ruby run.",
                    "ruby",
                    lossKind: HtmlConversionLossKind.Approximation);
            }
            return true;
        }

        private static Run CreateRubyContainer(string baseText, string rubyText, TextFormatting formatting, HtmlToWordOptions options) =>
            new(new Ruby(
                new RubyProperties(
                    new RubyAlign { Val = RubyAlignValues.Center },
                    new PhoneticGuideTextFontSize { Val = "10" },
                    new PhoneticGuideRaise { Val = 18 },
                    new PhoneticGuideBaseTextSize { Val = "20" },
                    new LanguageId { Val = string.IsNullOrWhiteSpace(formatting.Language) ? "en-US" : formatting.Language! }),
                new RubyContent(CreateRubyRun(rubyText, formatting, options)),
                new RubyBase(CreateRubyRun(baseText, formatting, options))));

        private static bool TryExtractRubyPairs(IElement element, out List<(string BaseText, string RubyText)> pairs, out bool approximated) {
            pairs = new List<(string BaseText, string RubyText)>();
            approximated = false;
            var explicitBases = element.Children
                .Where(child => child.TagName.Equals("rb", StringComparison.OrdinalIgnoreCase))
                .Select(child => NormalizeRubyText(child.TextContent))
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();
            var explicitAnnotations = element.Children
                .Where(child => child.TagName.Equals("rt", StringComparison.OrdinalIgnoreCase))
                .Select(child => NormalizeRubyText(child.TextContent))
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();
            if (explicitBases.Count > 0 && explicitBases.Count == explicitAnnotations.Count) {
                for (int index = 0; index < explicitBases.Count; index++) {
                    pairs.Add((explicitBases[index], explicitAnnotations[index]));
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

            pairs.Add((baseText, rubyText));
            approximated = explicitBases.Count > 1 || explicitAnnotations.Count > 1 ||
                           explicitBases.Count != explicitAnnotations.Count;
            return true;
        }

        private static string NormalizeRubyText(string text) =>
            string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

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
            if (formatting.FontSize.HasValue) properties.AppendChild(new FontSize { Val = (formatting.FontSize.Value * 2).ToString(CultureInfo.InvariantCulture) });
            if (formatting.Caps == CapsStyle.SmallCaps) properties.AppendChild(new SmallCaps());
            if (formatting.Caps == CapsStyle.Caps) properties.AppendChild(new Caps());
            if (formatting.LetterSpacing.HasValue) properties.AppendChild(new Spacing { Val = formatting.LetterSpacing.Value });
            if (!string.IsNullOrWhiteSpace(formatting.Language)) properties.AppendChild(new Languages { Val = formatting.Language! });

            var font = ResolveFontFamily(formatting.FontFamily) ?? ResolveFontFamily(options.FontFamily);
            if (!string.IsNullOrEmpty(font)) {
                properties.AppendChild(new RunFonts { Ascii = font, HighAnsi = font, EastAsia = font, ComplexScript = font });
            }

            return properties;
        }
    }
}
