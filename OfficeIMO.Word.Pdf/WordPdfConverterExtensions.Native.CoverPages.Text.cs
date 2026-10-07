using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static IReadOnlyList<PdfCore.PdfTextRun> GetNativeVmlTextRuns(WordDocument document, OpenXmlElement element) {
            var runs = new List<PdfCore.PdfTextRun>();
            bool hasParagraph = false;
            foreach (W.Paragraph paragraph in element.Descendants<W.Paragraph>()) {
                if (hasParagraph) {
                    runs.Add(PdfCore.PdfTextRun.LineBreak());
                }

                hasParagraph = true;
                foreach (W.Run run in paragraph.Descendants<W.Run>()) {
                    W.RunProperties? properties = run.RunProperties;
                    NativeTextSpacing spacing = ResolveNativeTextRunStyle(
                        new WordParagraph(document, paragraph, run)).TextSpacing;
                    foreach (OpenXmlElement child in run.ChildElements) {
                        if (child is W.Text text) {
                            AddNativeVmlTextRun(runs, document, run, properties, text, spacing);
                        } else if (child is W.Break) {
                            runs.Add(CreateNativeVmlTextRun("\n", properties, spacing));
                        } else if (child is W.TabChar) {
                            runs.Add(spacing.ApplyTo(PdfCore.PdfTextRun.Tab()));
                        } else {
                            foreach (W.Text nestedText in child.Descendants<W.Text>()) {
                                AddNativeVmlTextRun(runs, document, run, properties, nestedText, spacing);
                            }
                        }
                    }
                }
            }

            AddNativeVmlTextPathRuns(runs, document, element);

            while (runs.Count > 0 && runs[0].Text == "\n") {
                runs.RemoveAt(0);
            }

            while (runs.Count > 0 && runs[runs.Count - 1].Text == "\n") {
                runs.RemoveAt(runs.Count - 1);
            }

            return runs;
        }

        private static void AddNativeVmlTextRun(List<PdfCore.PdfTextRun> runs, WordDocument document, W.Run run, W.RunProperties? properties, W.Text text, NativeTextSpacing spacing) {
            string value = ResolveNativeVmlRunText(document, run, text);
            if (string.IsNullOrEmpty(value)) {
                return;
            }

            runs.Add(CreateNativeVmlTextRun(value, properties, spacing));
        }

        private static PdfCore.PdfTextRun CreateNativeVmlTextRun(string value, W.RunProperties? properties, NativeTextSpacing spacing) =>
            spacing.ApplyTo(new PdfCore.PdfTextRun(
                value,
                bold: HasNativeOnOff(properties?.Bold),
                underline: HasNativeVmlUnderline(properties?.Underline),
                color: ParseNativeColor(properties?.Color?.Val?.Value),
                italic: HasNativeOnOff(properties?.Italic),
                strike: HasNativeOnOff(properties?.Strike),
                fontSize: GetNativeVmlRunFontSize(properties),
                font: GetNativeVmlRunFont(properties)));

        private static void AddNativeVmlTextPathRuns(List<PdfCore.PdfTextRun> runs, WordDocument document, OpenXmlElement element) {
            foreach (V.TextPath textPath in element.Descendants<V.TextPath>()) {
                if (textPath.On != null && textPath.On.Value == false) {
                    continue;
                }

                string? raw = textPath.String?.Value;
                if (string.IsNullOrWhiteSpace(raw)) {
                    continue;
                }

                string value = ResolveNativeBuiltInPropertyPlaceholders(document, raw!);
                if (runs.Count > 0 && runs[runs.Count - 1].Text != "\n") {
                    runs.Add(PdfCore.PdfTextRun.LineBreak());
                }

                runs.Add(new PdfCore.PdfTextRun(
                    value,
                    bold: false,
                    underline: false,
                    color: null,
                    italic: false,
                    strike: false,
                    fontSize: GetNativeVmlTextPathFontSize(textPath),
                    font: GetNativeVmlTextPathFont(textPath)));
            }
        }

        private static double? GetNativeVmlTextPathFontSize(V.TextPath textPath) {
            Dictionary<string, string> style = ParseNativeVmlStyle(textPath.Style?.Value);
            double? fontSize = ResolveNativeVmlLength(style.TryGetValue("font-size", out string? value) ? value : null, 1D, 1D);
            return fontSize.HasValue && fontSize.Value > 0D && fontSize.Value <= MaxNativeVmlTextPathFontSizePoints
                ? fontSize
                : null;
        }

        private static PdfCore.PdfStandardFont? GetNativeVmlTextPathFont(V.TextPath textPath) {
            Dictionary<string, string> style = ParseNativeVmlStyle(textPath.Style?.Value);
            if (!style.TryGetValue("font-family", out string? family)) {
                return null;
            }

            family = family.Trim().Trim('"', '\'');
            return PdfCore.PdfStandardFontMapper.TryMapFontFamily(family, out PdfCore.PdfStandardFont font)
                ? font
                : null;
        }

        private static string ResolveNativeVmlRunText(WordDocument document, W.Run run, W.Text textElement) {
            string text = textElement.Text;
            string value = ResolveNativeBuiltInPropertyPlaceholders(document, text);
            string? propertyValue = GetNativeBuiltInPropertyValue(document, GetNativeVmlRunSdtProperties(textElement, run));
            if (!string.IsNullOrWhiteSpace(propertyValue) &&
                (string.Equals(value, text, StringComparison.Ordinal) || IsNativeVmlPlaceholderText(value))) {
                value = PreserveNativeVmlPlaceholderSpacing(text, propertyValue!);
            }

            W.RunProperties? properties = run.RunProperties;
            if (HasNativeOnOff(properties?.Caps) || HasNativeOnOff(properties?.SmallCaps)) {
                value = value.ToUpperInvariant();
            }

            return value;
        }

        private static W.SdtProperties? GetNativeVmlRunSdtProperties(W.Text textElement, W.Run run) {
            W.SdtRun? runSdt = textElement.Ancestors<W.SdtRun>().FirstOrDefault() ?? run.Ancestors<W.SdtRun>().FirstOrDefault();
            if (runSdt?.SdtProperties != null) {
                return runSdt.SdtProperties;
            }

            W.SdtBlock? blockSdt = textElement.Ancestors<W.SdtBlock>().FirstOrDefault() ?? run.Ancestors<W.SdtBlock>().FirstOrDefault();
            return blockSdt?.SdtProperties;
        }

        private static bool IsNativeVmlPlaceholderText(string text) {
            string trimmed = text.Trim();
            return trimmed.Length > 2 && trimmed[0] == '[' && trimmed[trimmed.Length - 1] == ']';
        }

        private static string PreserveNativeVmlPlaceholderSpacing(string sourceText, string value) {
            int leading = 0;
            while (leading < sourceText.Length && char.IsWhiteSpace(sourceText[leading])) {
                leading++;
            }

            int trailing = 0;
            while (trailing < sourceText.Length - leading && char.IsWhiteSpace(sourceText[sourceText.Length - trailing - 1])) {
                trailing++;
            }

            return sourceText.Substring(0, leading) + value + sourceText.Substring(sourceText.Length - trailing, trailing);
        }

        private static PdfCore.PdfCanvasTextBoxStyle CreateNativeVmlTextBoxStyle(OpenXmlElement element, IReadOnlyList<PdfCore.PdfTextRun> runs) {
            var style = new PdfCore.PdfCanvasTextBoxStyle {
                Background = null,
                BorderColor = null,
                BorderWidth = 0D,
                TextColor = null,
                Align = GetNativeVmlTextAlign(element),
                VerticalAlign = GetNativeVmlVerticalAlign(element),
                FontSize = GetNativeVmlDefaultFontSize(runs),
                LineHeight = GetNativeVmlDefaultFontSize(runs) * 1.2D
            };

            ApplyNativeVmlTextBoxInset(style, element);
            return style;
        }

        private static bool HasNativeOnOff(OpenXmlElement? element) {
            if (element == null) {
                return false;
            }

            string? value = GetNativeOpenXmlAttribute(element, "val");
            return value == null ||
                   value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("true", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasNativeVmlUnderline(W.Underline? underline) {
            if (underline == null) {
                return false;
            }

            if (underline.Val != null && underline.Val.Value == W.UnderlineValues.None) {
                return false;
            }

            string value = underline.Val?.Value.ToString() ?? string.Empty;
            if (value.Length == 0) {
                return true;
            }

            return !value.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                   !value.Equals("0", StringComparison.OrdinalIgnoreCase) &&
                   !value.Equals("false", StringComparison.OrdinalIgnoreCase);
        }

        private static double? GetNativeVmlRunFontSize(W.RunProperties? properties) {
            string? value = properties?.FontSize?.Val?.Value;
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints) && halfPoints > 0D
                ? halfPoints / 2D
                : null;
        }

        private static PdfCore.PdfStandardFont? GetNativeVmlRunFont(W.RunProperties? properties) {
            W.RunFonts? fonts = properties?.RunFonts;
            if (fonts == null) {
                return null;
            }

            foreach (string? family in new[] {
                fonts.Ascii?.Value,
                fonts.HighAnsi?.Value,
                fonts.EastAsia?.Value,
                fonts.ComplexScript?.Value
            }) {
                if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(family, out PdfCore.PdfStandardFont font)) {
                    return font;
                }
            }

            return null;
        }

        private static double GetNativeVmlDefaultFontSize(IReadOnlyList<PdfCore.PdfTextRun> runs) {
            double max = 0D;
            foreach (PdfCore.PdfTextRun run in runs) {
                if (!string.IsNullOrWhiteSpace(run.Text) && run.FontSize.HasValue && run.FontSize.Value > max) {
                    max = run.FontSize.Value;
                }
            }

            return max > 0D ? max : 11D;
        }

        private static PdfCore.PdfAlign GetNativeVmlTextAlign(OpenXmlElement element) {
            W.Justification? justification = element.Descendants<W.Justification>().FirstOrDefault();
            W.JustificationValues? value = justification?.Val?.Value;
            if (value == W.JustificationValues.Center) return PdfCore.PdfAlign.Center;
            if (value == W.JustificationValues.Right) return PdfCore.PdfAlign.Right;
            return PdfCore.PdfAlign.Left;
        }

        private static PdfCore.PdfVerticalAlign GetNativeVmlVerticalAlign(OpenXmlElement element) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            if (style.TryGetValue("v-text-anchor", out string? anchor)) {
                if (anchor.Equals("middle", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfVerticalAlign.Middle;
                if (anchor.Equals("bottom", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfVerticalAlign.Bottom;
            }

            return PdfCore.PdfVerticalAlign.Top;
        }

        private static void ApplyNativeVmlTextBoxInset(PdfCore.PdfCanvasTextBoxStyle style, OpenXmlElement element) {
            OpenXmlElement? textBox = element.Descendants().FirstOrDefault(child => child.NamespaceUri == "urn:schemas-microsoft-com:vml" && child.LocalName == "textbox");
            string? inset = textBox == null ? null : GetNativeOpenXmlAttribute(textBox, "inset");
            if (string.IsNullOrWhiteSpace(inset)) {
                style.PaddingX = 0D;
                style.PaddingY = 0D;
                return;
            }

            string[] parts = inset!.Split(',');
            if (parts.Length < 2) {
                return;
            }

            style.PaddingLeft = ResolveNativeVmlLength(parts[0], 0D, 0D) ?? 0D;
            style.PaddingTop = ResolveNativeVmlLength(parts[1], 0D, 0D) ?? 0D;
            style.PaddingRight = parts.Length > 2 ? ResolveNativeVmlLength(parts[2], 0D, 0D) ?? style.PaddingLeft : style.PaddingLeft;
            style.PaddingBottom = parts.Length > 3 ? ResolveNativeVmlLength(parts[3], 0D, 0D) ?? style.PaddingTop : style.PaddingTop;
        }

    }
}
