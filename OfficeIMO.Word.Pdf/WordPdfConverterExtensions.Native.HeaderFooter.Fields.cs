using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private static bool TryBuildNativeHeaderFooterParagraphText(WordParagraph paragraph, out string? text, out PdfCore.PdfPageNumberStyle? pageNumberStyle, List<(OfficeXmlNode Run, string Text, bool IsField, OfficeXmlNode SourceChild, PdfCore.PdfPageNumberStyle? FieldStyle)>? serializedRuns = null) {
            text = null;
            pageNumberStyle = null;
            OfficeXmlNode source = paragraph.Node();
            if (source.IsNone) {
                return false;
            }

            var builder = new StringBuilder();
            WordComplexFieldRunVisibility prefix = WordConverterNodes.FieldVisibilityForParagraph(source);
            var state = new NativeHeaderFooterFieldState {
                Paragraph = paragraph,
                SerializedRuns = serializedRuns,
                CollectingFieldCode = prefix.HasOpenField && !prefix.IsVisible,
                SkippingFieldResult = prefix.HasOpenField && prefix.IsVisible
            };
            string prefixCode = ReadNativeHeaderFooterFieldPrefixCode(prefix.BeginMarkerNode(), source);
            if (state.CollectingFieldCode) state.FieldCode.Append(prefixCode);
            bool hasFieldToken = state.SkippingFieldResult && TryGetNativeHeaderFooterFieldToken(prefixCode, out _, out _);
            bool hasConflictingStyles = false;
            foreach (OfficeXmlNode element in source.Elements()) {
                AppendNativeHeaderFooterElementText(element, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
            }

            if (!hasFieldToken) {
                pageNumberStyle = null;
                return false;
            }

            text = builder.ToString();
            // A recognized but hidden field must not reach the field-only fallback.
            return true;
        }

        private static void AppendNativeHeaderFooterElementText(OfficeXmlNode element, StringBuilder builder, NativeHeaderFooterFieldState state, ref PdfCore.PdfPageNumberStyle? pageNumberStyle, ref bool hasConflictingStyles, ref bool hasFieldToken) {
            if (element.IsW("r")) {
                AppendNativeHeaderFooterRunText(element, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                return;
            }

            if (element.IsW("hyperlink")) {
                foreach (OfficeXmlNode childRun in element.WElements("r")) {
                    AppendNativeHeaderFooterRunText(childRun, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }

                return;
            }

            if (element.IsW("sdt")) {
                foreach (OfficeXmlNode child in element.WChild("sdtContent").Elements()) {
                    AppendNativeHeaderFooterElementText(child, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }

                return;
            }

            if (element.IsW("fldSimple")) {
                if (state.CollectingFieldCode || state.SkippingFieldResult) return;
                string fieldCode = element.WAttr("instr") ?? string.Empty;
                if (TryGetNativeHeaderFooterFieldToken(fieldCode, out string? token, out PdfCore.PdfPageNumberStyle? style)) {
                    hasFieldToken = true;
                    OfficeXmlNode[] resultRuns = element.Descendants(WordXml.W, "r").Where(run => run.Child(WordXml.W, "t").Exists).ToArray();
                    OfficeXmlNode resultRun = resultRuns.FirstOrDefault(run => !IsNativeHiddenHeaderFooterRun(run, state.Paragraph));
                    if (resultRun.Exists) {
                        AppendNativeHeaderFooterVisibleFieldToken(builder, resultRun, token!, style, state, ref pageNumberStyle, ref hasConflictingStyles);
                    } else if (resultRuns.Length == 0) {
                        // An uncached field still calculates a value using its effective formatting.
                        OfficeXmlNode emptyRun = element.FirstDescendantElement(WordXml.W, "r");
                        bool hidden = IsNativeHiddenHeaderFooterRun(emptyRun.Exists ? emptyRun : WordConverterNodes.CreateDetachedRun(), state.Paragraph);
                        if (!hidden) AppendNativeHeaderFooterVisibleFieldToken(builder, emptyRun, token!, style, state, ref pageNumberStyle, ref hasConflictingStyles);
                    }
                    return;
                }

                foreach (OfficeXmlNode child in element.Elements()) {
                    AppendNativeHeaderFooterElementText(child, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }
            }
        }

        private static void AppendNativeHeaderFooterRunText(OfficeXmlNode run, StringBuilder builder, NativeHeaderFooterFieldState state, ref PdfCore.PdfPageNumberStyle? pageNumberStyle, ref bool hasConflictingStyles, ref bool hasFieldToken) {
            WordParagraph sourceRun = state.Paragraph._document.RunView(state.Paragraph.Node(), run);
            bool hidden = IsNativeHiddenTextRun(sourceRun, state.Paragraph);

            foreach (OfficeXmlNode child in run.Elements()) {
                if (child.IsW("fldChar")) {
                    string? fieldCharType = child.WAttr("fldCharType");
                    if (fieldCharType == "begin") {
                        state.CollectingFieldCode = true;
                        state.SkippingFieldResult = false;
                        state.FieldCode.Clear();
                    } else if (fieldCharType == "separate") {
                        if (TryGetNativeHeaderFooterFieldToken(state.FieldCode.ToString(), out string? token, out PdfCore.PdfPageNumberStyle? style)) {
                            bool visibleResult = TryGetVisibleNativeHeaderFooterFieldResult(child, state.Paragraph,
                                ref state.FieldScanWork, out OfficeXmlNode resultRun);
                            if (visibleResult) {
                                builder.Append(token);
                                MergeNativeHeaderFooterPageNumberStyle(ref pageNumberStyle, ref hasConflictingStyles, style);
                                state.SerializedRuns?.Add((resultRun, token!, true, default, style));
                            }
                            hasFieldToken = true;
                            state.SkippingFieldResult = true;
                        }

                        state.CollectingFieldCode = false;
                    } else if (fieldCharType == "end") {
                        // The separator is optional when a complex field has no cached result.
                        if (state.CollectingFieldCode && TryGetNativeHeaderFooterFieldToken(state.FieldCode.ToString(),
                            out string? token, out PdfCore.PdfPageNumberStyle? style)) {
                            hasFieldToken = true;
                            if (!hidden) AppendNativeHeaderFooterVisibleFieldToken(builder, run, token!, style, state,
                                ref pageNumberStyle, ref hasConflictingStyles);
                        }
                        state.CollectingFieldCode = false;
                        state.SkippingFieldResult = false;
                        state.FieldCode.Clear();
                    }

                    continue;
                }

                if (child.IsW("instrText")) {
                    if (state.CollectingFieldCode) {
                        state.FieldCode.Append(child.InnerText);
                    }

                    continue;
                }

                if (state.CollectingFieldCode || state.SkippingFieldResult) {
                    // A recognized field token is emitted at its separator. Cached text is skipped.
                    continue;
                }

                if (child.IsW("pgNum")) {
                    hasFieldToken = true;
                    if (!hidden) AppendNativeHeaderFooterVisibleFieldToken(builder, run, "{page}", null, state,
                        ref pageNumberStyle, ref hasConflictingStyles);
                    continue;
                }

                if (hidden) continue;
                if (child.IsW("t")) {
                    string sourceText = WordConverterNodes.VisibleText(child);
                    string visibleText = ApplyNativeTextTransform(sourceText, sourceRun, state.Paragraph);
                    builder.Append(visibleText);
                    // Keep source case until styled replacement segmentation has
                    // distinguished original capitals from synthetic small caps.
                    state.SerializedRuns?.Add((run, sourceText, false, child, null));
                } else if (child.IsW("tab") || child.IsW("ptab")) {
                    builder.Append('\t');
                    state.SerializedRuns?.Add((run, "\t", false, child, null));
                } else if (child.IsW("br")) {
                    builder.AppendLine();
                    state.SerializedRuns?.Add((run, Environment.NewLine, false, child, null));
                } else if (child.IsW("drawing") || child.IsW("pict") || child.Is(WordXml.Mc, "AlternateContent")) {
                    string drawingText = WordConverterNodes.VisibleText(child);
                    builder.Append(drawingText);
                    state.SerializedRuns?.Add((run, drawingText, false, child, null));
                }
            }
        }

        private static bool IsNativeHiddenHeaderFooterRun(OfficeXmlNode run, WordParagraph paragraph) =>
            IsNativeHiddenTextRun(paragraph._document.RunView(paragraph.Node(), run), paragraph);

        private static void AppendNativeHeaderFooterVisibleFieldToken(StringBuilder builder, OfficeXmlNode run, string token,
            PdfCore.PdfPageNumberStyle? style, NativeHeaderFooterFieldState state,
            ref PdfCore.PdfPageNumberStyle? pageNumberStyle, ref bool hasConflictingStyles) {
            builder.Append(token);
            state.SerializedRuns?.Add((run.Exists ? run : WordConverterNodes.CreateDetachedRun(), token, true, default, style));
            MergeNativeHeaderFooterPageNumberStyle(ref pageNumberStyle, ref hasConflictingStyles, style);
        }

        private static bool TryGetNativeHeaderFooterFieldToken(WordParagraph paragraph, out string? token, out PdfCore.PdfPageNumberStyle? style) {
            token = null;
            style = null;
            WordField? field = paragraph.Field;
            if (field?.FieldType == WordFieldType.Page) {
                token = "{page}";
                style = MapNativePageNumberFieldStyle(field.Field);
                return true;
            }

            if (field?.FieldType == WordFieldType.NumPages) {
                token = "{documentpages}";
                style = MapNativePageNumberFieldStyle(field.Field);
                return true;
            }

            if (field?.FieldType == WordFieldType.SectionPages) {
                token = "{sectionpages}";
                style = MapNativePageNumberFieldStyle(field.Field);
                return true;
            }

            return false;
        }

        private static bool TryGetNativeHeaderFooterFieldToken(string fieldCode, out string? token, out PdfCore.PdfPageNumberStyle? style) {
            token = null;
            style = null;
            string trimmed = fieldCode.Trim();
            if (trimmed.Length == 0) {
                return false;
            }

            int end = 0;
            while (end < trimmed.Length && !char.IsWhiteSpace(trimmed[end])) {
                end++;
            }

            string fieldType = trimmed.Substring(0, end);
            if (string.Equals(fieldType, "PAGE", StringComparison.OrdinalIgnoreCase)) {
                token = "{page}";
                style = MapNativePageNumberFieldStyle(trimmed);
                return true;
            }

            if (string.Equals(fieldType, "NUMPAGES", StringComparison.OrdinalIgnoreCase)) {
                token = "{documentpages}";
                style = MapNativePageNumberFieldStyle(trimmed);
                return true;
            }

            if (string.Equals(fieldType, "SECTIONPAGES", StringComparison.OrdinalIgnoreCase)) {
                token = "{sectionpages}";
                style = MapNativePageNumberFieldStyle(trimmed);
                return true;
            }

            return false;
        }

        private static PdfCore.PdfPageNumberStyle? MapNativePageNumberFieldStyle(string fieldCode) {
            string? format = GetNativePageNumberFieldFormatSwitch(fieldCode);
            if (format == "roman") {
                return PdfCore.PdfPageNumberStyle.LowerRoman;
            }

            if (format == "Roman" || format == "ROMAN") {
                return PdfCore.PdfPageNumberStyle.UpperRoman;
            }

            if (format == "Alphabetical") {
                return PdfCore.PdfPageNumberStyle.LowerLetter;
            }

            if (format == "ALPHABETICAL") {
                return PdfCore.PdfPageNumberStyle.UpperLetter;
            }

            if (format == "Arabic") {
                return PdfCore.PdfPageNumberStyle.Arabic;
            }

            return null;
        }

        private static string? GetNativePageNumberFieldFormatSwitch(string fieldCode) {
            int markerIndex = fieldCode.IndexOf(@"\*", StringComparison.Ordinal);
            while (markerIndex >= 0) {
                int index = markerIndex + 2;
                while (index < fieldCode.Length && char.IsWhiteSpace(fieldCode[index])) {
                    index++;
                }

                int start = index;
                while (index < fieldCode.Length && (char.IsLetter(fieldCode[index]) || fieldCode[index] == '_')) {
                    index++;
                }

                if (index > start) {
                    string format = fieldCode.Substring(start, index - start);
                    if (!string.Equals(format, "MERGEFORMAT", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(format, "CHARFORMAT", StringComparison.OrdinalIgnoreCase)) {
                        return format;
                    }
                }

                markerIndex = fieldCode.IndexOf(@"\*", markerIndex + 2, StringComparison.Ordinal);
            }

            return null;
        }

        private static void MergeNativeHeaderFooterPageNumberStyle(ref PdfCore.PdfPageNumberStyle? current, ref bool hasConflict, PdfCore.PdfPageNumberStyle? candidate) {
            if (!candidate.HasValue || hasConflict) {
                return;
            }

            if (current.HasValue && current.Value != candidate.Value) {
                current = null;
                hasConflict = true;
                return;
            }

            current = candidate.Value;
        }

        private sealed class NativeHeaderFooterFieldState {
            public WordParagraph Paragraph { get; set; } = null!;
            public bool CollectingFieldCode { get; set; }
            public bool SkippingFieldResult { get; set; }
            public int FieldScanWork;
            public List<(OfficeXmlNode Run, string Text, bool IsField, OfficeXmlNode SourceChild, PdfCore.PdfPageNumberStyle? FieldStyle)>? SerializedRuns { get; set; }
            public StringBuilder FieldCode { get; } = new StringBuilder();
        }

    }
}
