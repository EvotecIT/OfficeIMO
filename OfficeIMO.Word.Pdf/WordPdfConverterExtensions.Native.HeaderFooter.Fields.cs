using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool TryBuildNativeHeaderFooterParagraphText(WordParagraph paragraph, out string? text, out PdfCore.PdfPageNumberStyle? pageNumberStyle, List<(W.Run Run, string Text, bool IsField, DocumentFormat.OpenXml.OpenXmlElement? SourceChild)>? serializedRuns = null) {
            text = null;
            pageNumberStyle = null;
            if (paragraph._paragraph == null) {
                return false;
            }

            var builder = new StringBuilder();
            WordComplexFieldRunVisibility prefix = WordComplexFieldRunVisibility.ForParagraph(paragraph._paragraph);
            var state = new NativeHeaderFooterFieldState {
                Paragraph = paragraph,
                SerializedRuns = serializedRuns,
                CollectingFieldCode = prefix.HasOpenField && !prefix.IsVisible,
                SkippingFieldResult = prefix.HasOpenField && prefix.IsVisible
            };
            string prefixCode = ReadNativeHeaderFooterFieldPrefixCode(prefix.CurrentState?.BeginMarker, paragraph._paragraph);
            if (state.CollectingFieldCode) state.FieldCode.Append(prefixCode);
            bool hasFieldToken = state.SkippingFieldResult && TryGetNativeHeaderFooterFieldToken(prefixCode, out _, out _);
            bool hasConflictingStyles = false;
            foreach (var element in paragraph._paragraph.ChildElements) {
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

        private static void AppendNativeHeaderFooterElementText(DocumentFormat.OpenXml.OpenXmlElement element, StringBuilder builder, NativeHeaderFooterFieldState state, ref PdfCore.PdfPageNumberStyle? pageNumberStyle, ref bool hasConflictingStyles, ref bool hasFieldToken) {
            if (element is W.Run run) {
                AppendNativeHeaderFooterRunText(run, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                return;
            }

            if (element is W.Hyperlink hyperlink) {
                foreach (W.Run childRun in hyperlink.Elements<W.Run>()) {
                    AppendNativeHeaderFooterRunText(childRun, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }

                return;
            }

            if (element is W.SdtRun sdtRun) {
                foreach (var child in sdtRun.SdtContentRun?.ChildElements ?? Enumerable.Empty<DocumentFormat.OpenXml.OpenXmlElement>()) {
                    AppendNativeHeaderFooterElementText(child, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }

                return;
            }

            if (element is W.SimpleField simpleField) {
                if (state.CollectingFieldCode || state.SkippingFieldResult) return;
                string fieldCode = simpleField.Instruction?.Value ?? string.Empty;
                if (TryGetNativeHeaderFooterFieldToken(fieldCode, out string? token, out PdfCore.PdfPageNumberStyle? style)) {
                    hasFieldToken = true;
                    W.Run[] resultRuns = simpleField.Descendants<W.Run>().Where(run => run.Elements<W.Text>().Any()).ToArray();
                    W.Run? resultRun = resultRuns.FirstOrDefault(run => !IsNativeHiddenHeaderFooterRun(run, state.Paragraph));
                    if (resultRun != null) {
                        AppendNativeHeaderFooterVisibleFieldToken(builder, resultRun, token!, style, state, ref pageNumberStyle, ref hasConflictingStyles);
                    } else if (resultRuns.Length == 0) {
                        // An uncached field still calculates a value using its effective formatting.
                        W.Run? emptyRun = simpleField.Descendants<W.Run>().FirstOrDefault();
                        bool hidden = IsNativeHiddenHeaderFooterRun(emptyRun ?? new W.Run(), state.Paragraph);
                        if (!hidden) AppendNativeHeaderFooterVisibleFieldToken(builder, emptyRun, token!, style, state, ref pageNumberStyle, ref hasConflictingStyles);
                    }
                    return;
                }

                foreach (var child in simpleField.ChildElements) {
                    AppendNativeHeaderFooterElementText(child, builder, state, ref pageNumberStyle, ref hasConflictingStyles, ref hasFieldToken);
                }
            }
        }

        private static void AppendNativeHeaderFooterRunText(W.Run run, StringBuilder builder, NativeHeaderFooterFieldState state, ref PdfCore.PdfPageNumberStyle? pageNumberStyle, ref bool hasConflictingStyles, ref bool hasFieldToken) {
            var sourceRun = new WordParagraph(state.Paragraph._document, state.Paragraph._paragraph!, run);
            bool hidden = IsNativeHiddenTextRun(sourceRun, state.Paragraph);

            foreach (var child in run.ChildElements) {
                if (child is W.FieldChar fieldChar) {
                    W.FieldCharValues? fieldCharType = fieldChar.FieldCharType?.Value;
                    if (fieldCharType == W.FieldCharValues.Begin) {
                        state.CollectingFieldCode = true;
                        state.SkippingFieldResult = false;
                        state.FieldCode.Clear();
                        state.PendingToken = null;
                    } else if (fieldCharType == W.FieldCharValues.Separate) {
                        if (TryGetNativeHeaderFooterFieldToken(state.FieldCode.ToString(), out string? token, out PdfCore.PdfPageNumberStyle? style)) {
                            bool visibleResult = HasVisibleNativeHeaderFooterFieldResult(fieldChar, state.Paragraph);
                            state.PendingToken = visibleResult ? token : null;
                            if (visibleResult) {
                                builder.Append(token);
                                MergeNativeHeaderFooterPageNumberStyle(ref pageNumberStyle, ref hasConflictingStyles, style);
                            }
                            hasFieldToken = true;
                            state.SkippingFieldResult = true;
                        }

                        state.CollectingFieldCode = false;
                    } else if (fieldCharType == W.FieldCharValues.End) {
                        state.CollectingFieldCode = false;
                        state.SkippingFieldResult = false;
                        state.FieldCode.Clear();
                        state.PendingToken = null;
                    }

                    continue;
                }

                if (child is W.FieldCode fieldCode) {
                    if (state.CollectingFieldCode) {
                        state.FieldCode.Append(fieldCode.Text);
                    }

                    continue;
                }

                if (state.CollectingFieldCode || state.SkippingFieldResult) {
                    if (state.SkippingFieldResult && child is W.Text) {
                        if (state.PendingToken != null && !hidden) {
                            state.SerializedRuns?.Add((run, state.PendingToken, true, null));
                            state.PendingToken = null;
                        }
                    }
                    continue;
                }

                if (hidden) continue;
                if (child is W.Text text) {
                    string visibleText = ApplyNativeTextTransform(WordParagraph.ReadVisibleText(text), sourceRun, state.Paragraph);
                    builder.Append(visibleText);
                    state.SerializedRuns?.Add((run, visibleText, false, child));
                } else if (child is W.TabChar) {
                    builder.Append('\t');
                    state.SerializedRuns?.Add((run, "\t", false, child));
                } else if (child is W.Break) {
                    builder.AppendLine();
                    state.SerializedRuns?.Add((run, Environment.NewLine, false, child));
                } else if (child is W.Drawing or W.Picture or DocumentFormat.OpenXml.AlternateContent) {
                    string drawingText = WordParagraph.ReadVisibleText(child);
                    builder.Append(drawingText);
                    state.SerializedRuns?.Add((run, drawingText, false, child));
                }
            }
        }

        private static bool IsNativeHiddenHeaderFooterRun(W.Run run, WordParagraph paragraph) =>
            IsNativeHiddenTextRun(new WordParagraph(paragraph._document, paragraph._paragraph!, run), paragraph);

        private static void AppendNativeHeaderFooterVisibleFieldToken(StringBuilder builder, W.Run? run, string token,
            PdfCore.PdfPageNumberStyle? style, NativeHeaderFooterFieldState state,
            ref PdfCore.PdfPageNumberStyle? pageNumberStyle, ref bool hasConflictingStyles) {
            builder.Append(token);
            if (run != null) state.SerializedRuns?.Add((run, token, true, null));
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
                token = "{pages}";
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
                token = "{pages}";
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

            if (format == "Roman") {
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
                    return fieldCode.Substring(start, index - start);
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
            public List<(W.Run Run, string Text, bool IsField, DocumentFormat.OpenXml.OpenXmlElement? SourceChild)>? SerializedRuns { get; set; }
            public string? PendingToken { get; set; }
            public StringBuilder FieldCode { get; } = new StringBuilder();
        }

    }
}
