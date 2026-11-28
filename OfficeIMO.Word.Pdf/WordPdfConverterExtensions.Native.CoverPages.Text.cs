using System.Globalization;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private static IReadOnlyList<PdfCore.PdfTextRun> GetNativeVmlTextRuns(WordDocument document, OfficeXmlNode element) {
            var runs = new List<PdfCore.PdfTextRun>();
            bool hasParagraph = false;
            foreach (OfficeXmlNode paragraph in element.ModelDescendants("p")) {
                if (hasParagraph) {
                    runs.Add(PdfCore.PdfTextRun.LineBreak());
                }

                hasParagraph = true;
                foreach (OfficeXmlNode run in paragraph.ModelDescendants("r")) {
                    OfficeXmlNode properties = run.WChild("rPr");
                    NativeResolvedTextStyle style = ResolveNativeTextRunStyle(
                        WordConverterNodes.RunView(document, paragraph, run));
                    foreach (OfficeXmlNode child in run.Elements()) {
                        if (child.IsW("t")) {
                            AddNativeVmlTextRun(runs, document, run, properties, child, style);
                        } else if (child.IsW("br")) {
                            runs.Add(CreateNativeVmlTextRun("\n", properties, style));
                        } else if (child.IsW("tab")) {
                            runs.Add(style.ApplyTo(PdfCore.PdfTextRun.Tab()));
                        } else {
                            foreach (OfficeXmlNode nestedText in child.Descendants(WordXml.W, "t")) {
                                AddNativeVmlTextRun(runs, document, run, properties, nestedText, style);
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

        private static void AddNativeVmlTextRun(List<PdfCore.PdfTextRun> runs, WordDocument document, OfficeXmlNode run, OfficeXmlNode properties, OfficeXmlNode text, NativeResolvedTextStyle style) {
            string value = ResolveNativeVmlRunText(document, run, text);
            if (string.IsNullOrEmpty(value)) {
                return;
            }

            foreach (NativeCasedTextSegment segment in GetNativeCasedTextSegments(value, style))
                runs.Add(CreateNativeVmlTextRun(segment.Text, properties, segment.Style));
        }

        private static PdfCore.PdfTextRun CreateNativeVmlTextRun(string value, OfficeXmlNode properties, NativeResolvedTextStyle style) =>
            style.ApplyTo(new PdfCore.PdfTextRun(
                value,
                bold: HasNativeOnOff(properties.WChild("b")),
                underline: HasNativeVmlUnderline(properties.WChild("u")),
                color: ParseNativeColor(properties.WChild("color").WVal()),
                italic: HasNativeOnOff(properties.WChild("i")),
                strike: HasNativeOnOff(properties.WChild("strike")),
                fontSize: ((double?)(style.FontSize ?? (style.SmallCaps && !style.AllCaps
                    ? style.EffectiveFontSize : GetNativeVmlRunFontSize(properties))) is double physicalLength1 ? global::OfficeIMO.OfficeLength.Points(physicalLength1) : (global::OfficeIMO.OfficeLength?)null),
                standardFont: GetNativeVmlRunFont(properties)));

        private static void AddNativeVmlTextPathRuns(List<PdfCore.PdfTextRun> runs, WordDocument document, OfficeXmlNode element) {
            foreach (OfficeXmlNode textPath in element.Descendants(WordXml.V, "textpath")) {
                if (textPath.Attr("on") is "f" or "false") {
                    continue;
                }

                string? raw = textPath.Attr("string");
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
                    fontSize: ((double?)(GetNativeVmlTextPathFontSize(textPath)) is double physicalLength2 ? global::OfficeIMO.OfficeLength.Points(physicalLength2) : (global::OfficeIMO.OfficeLength?)null),
                    standardFont: GetNativeVmlTextPathFont(textPath)));
            }
        }

        private static double? GetNativeVmlTextPathFontSize(OfficeXmlNode textPath) {
            Dictionary<string, string> style = ParseNativeVmlStyle(textPath.Attr("style"));
            double? fontSize = ResolveNativeVmlLength(style.TryGetValue("font-size", out string? value) ? value : null, 1D, 1D);
            return fontSize.HasValue && fontSize.Value > 0D && fontSize.Value <= MaxNativeVmlTextPathFontSizePoints
                ? fontSize
                : null;
        }

        private static PdfCore.PdfStandardFont? GetNativeVmlTextPathFont(OfficeXmlNode textPath) {
            Dictionary<string, string> style = ParseNativeVmlStyle(textPath.Attr("style"));
            if (!style.TryGetValue("font-family", out string? family)) {
                return null;
            }

            family = family.Trim();
            return PdfCore.PdfStandardFontMapper.TryMapFontFamily(family, out PdfCore.PdfStandardFont font)
                ? font
                : null;
        }

        private static string ResolveNativeVmlRunText(WordDocument document, OfficeXmlNode run, OfficeXmlNode textElement) {
            string text = textElement.VisibleText();
            string value = ResolveNativeBuiltInPropertyPlaceholders(document, text);
            string? propertyValue = GetNativeBuiltInPropertyValue(document, GetNativeVmlRunSdtProperties(textElement, run));
            if (!string.IsNullOrWhiteSpace(propertyValue) &&
                (string.Equals(value, text, StringComparison.Ordinal) || IsNativeVmlPlaceholderText(value))) {
                value = PreserveNativeVmlPlaceholderSpacing(text, propertyValue!);
            }

            return value;
        }

        private static OfficeXmlNode GetNativeVmlRunSdtProperties(OfficeXmlNode textElement, OfficeXmlNode run) {
            OfficeXmlNode runSdt = FindNativeVmlContentControl(textElement, WordContentControlLevel.Run);
            if (runSdt.IsNone) runSdt = FindNativeVmlContentControl(run, WordContentControlLevel.Run);
            if (runSdt.WChild("sdtPr").Exists) {
                return runSdt.WChild("sdtPr");
            }

            OfficeXmlNode blockSdt = FindNativeVmlContentControl(textElement, WordContentControlLevel.Block);
            if (blockSdt.IsNone) blockSdt = FindNativeVmlContentControl(run, WordContentControlLevel.Block);
            return blockSdt.WChild("sdtPr");
        }

        private static OfficeXmlNode FindNativeVmlContentControl(OfficeXmlNode element, WordContentControlLevel level) {
            foreach (OfficeXmlNode ancestor in element.Ancestors()) {
                if (ancestor.ContentControlLevel() == level) return ancestor;
            }
            return default;
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

        private static PdfCore.PdfCanvasTextBoxStyle CreateNativeVmlTextBoxStyle(WordDocument document, OfficeXmlNode element, IReadOnlyList<PdfCore.PdfTextRun> runs) {
            var style = new PdfCore.PdfCanvasTextBoxStyle {
                Background = null,
                BorderColor = null,
                BorderWidth = global::OfficeIMO.OfficeLength.Points(0D),
                TextColor = null,
                Align = GetNativeVmlTextAlign(document, element),
                VerticalAlign = GetNativeVmlVerticalAlign(element),
                FontSize = global::OfficeIMO.OfficeLength.Points(GetNativeVmlDefaultFontSize(runs)),
                LineHeight = global::OfficeIMO.OfficeLength.Points(GetNativeVmlDefaultFontSize(runs) * 1.2D)
            };

            ApplyNativeVmlTextBoxInset(style, element);
            return style;
        }

        private static bool HasNativeOnOff(OfficeXmlNode element) {
            if (element.IsNone) {
                return false;
            }

            string? value = GetNativeOpenXmlAttribute(element, "val");
            return value == null ||
                   value.Equals("1", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("true", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasNativeVmlUnderline(OfficeXmlNode underline) {
            if (underline.IsNone) {
                return false;
            }

            if (underline.WVal() == "none") {
                return false;
            }

            string value = underline.WVal() ?? string.Empty;
            if (value.Length == 0) {
                return true;
            }

            return !value.Equals("none", StringComparison.OrdinalIgnoreCase) &&
                   !value.Equals("0", StringComparison.OrdinalIgnoreCase) &&
                   !value.Equals("false", StringComparison.OrdinalIgnoreCase);
        }

        private static double? GetNativeVmlRunFontSize(OfficeXmlNode properties) {
            string? value = properties.WChild("sz").WVal();
            return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double halfPoints) && halfPoints > 0D
                ? halfPoints / 2D
                : null;
        }

        private static PdfCore.PdfStandardFont? GetNativeVmlRunFont(OfficeXmlNode properties) {
            OfficeXmlNode fonts = properties.WChild("rFonts");
            if (fonts.IsNone) {
                return null;
            }

            foreach (string? family in new[] {
                fonts.WAttr("ascii"),
                fonts.WAttr("hAnsi"),
                fonts.WAttr("eastAsia"),
                fonts.WAttr("cs")
            }) {
                if (TryResolveNativeMappedFont(family, nativeFontMap: null, out PdfCore.PdfStandardFont font)) {
                    return font;
                }
            }

            return null;
        }

        private static double GetNativeVmlDefaultFontSize(IReadOnlyList<PdfCore.PdfTextRun> runs) {
            double max = 0D;
            foreach (PdfCore.PdfTextRun run in runs) {
                if (!string.IsNullOrWhiteSpace(run.Text) && run.FontSize is OfficeLength fontSize) {
                    max = Math.Max(max, fontSize.ToPoints());
                }
            }

            return max > 0D ? max : 11D;
        }

        private static global::OfficeIMO.OfficeAlignment GetNativeVmlTextAlign(WordDocument document, OfficeXmlNode element) {
            // Paragraph justification sits in a w:pPr; a w:jc in table or row properties is a different setting.
            OfficeXmlNode justification = element.DescendantElements().FirstOrDefault(node => node.IsW("jc") && node.Parent.IsW("pPr") && node.IsModelElement());
            OfficeXmlNode source = justification.Ancestor(WordXml.W, "p");
            if (source.IsNone) source = element.ModelDescendants("p").FirstOrDefault();
            // The frame takes its paragraph's resolved alignment and direction: direct, style and document defaults.
            string? value = source.IsNone ? justification.WVal()
                : ResolveNativeUnjustifiedJustification(WordConverterNodes.ParagraphView(document, source));
            if (value == "center") return global::OfficeIMO.OfficeAlignment.Center;
            if (value == "right") return global::OfficeIMO.OfficeAlignment.Right;
            return global::OfficeIMO.OfficeAlignment.Left;
        }

        private static PdfCore.PdfVerticalAlign GetNativeVmlVerticalAlign(OfficeXmlNode element) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            if (style.TryGetValue("v-text-anchor", out string? anchor)) {
                if (anchor.Equals("middle", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfVerticalAlign.Middle;
                if (anchor.Equals("bottom", StringComparison.OrdinalIgnoreCase)) return PdfCore.PdfVerticalAlign.Bottom;
            }

            return PdfCore.PdfVerticalAlign.Top;
        }

        private static void ApplyNativeVmlTextBoxInset(PdfCore.PdfCanvasTextBoxStyle style, OfficeXmlNode element) {
            OfficeXmlNode textBox = element.DescendantElements().FirstOrDefault(child => child.NamespaceUri == "urn:schemas-microsoft-com:vml" && child.LocalName == "textbox");
            string? inset = textBox.IsNone ? null : GetNativeOpenXmlAttribute(textBox, "inset");
            if (string.IsNullOrWhiteSpace(inset)) {
                style.PaddingX = global::OfficeIMO.OfficeLength.Points(0D);
                style.PaddingY = global::OfficeIMO.OfficeLength.Points(0D);
                return;
            }

            string[] parts = inset!.Split(',');
            if (parts.Length < 2) {
                return;
            }

            style.PaddingLeft = global::OfficeIMO.OfficeLength.Points(ResolveNativeVmlLength(parts[0], 0D, 0D) ?? 0D);
            style.PaddingTop = global::OfficeIMO.OfficeLength.Points(ResolveNativeVmlLength(parts[1], 0D, 0D) ?? 0D);
            style.PaddingRight = parts.Length > 2 && ResolveNativeVmlLength(parts[2], 0D, 0D) is double rightPadding ? global::OfficeIMO.OfficeLength.Points(rightPadding) : style.PaddingLeft;
            style.PaddingBottom = parts.Length > 3 && ResolveNativeVmlLength(parts[3], 0D, 0D) is double bottomPadding ? global::OfficeIMO.OfficeLength.Points(bottomPadding) : style.PaddingTop;
        }

    }
}
