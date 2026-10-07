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
        private const double MaximumNativeHeaderFooterListOffsetPoints = 14_400D;
        private const int MaximumNativeHeaderFooterListSpacingCharacters = 8_192;

        private static void AddNativeHeaderFooterParagraphText(NativeHeaderFooterText parts, WordParagraph paragraph, IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeHeaderFooterZone? forcedZone = null, NativeFontMap? nativeFontMap = null) {
            WordTextBox? textBox = GetNativeParagraphTextBox(paragraph, out _);
            string? text = GetNativeHeaderFooterParagraphText(paragraph, listMarkers, out PdfCore.PdfPageNumberStyle? pageNumberStyle, out NativeHeaderFooterZone? zoneOverride);
            IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements = textBox == null
                ? CreateNativeHeaderFooterSpacingReplacements(paragraph, nativeFontMap)
                : CreateNativeHeaderFooterTextBoxListReplacements(textBox, listMarkers, nativeFontMap);
            AddNativeHeaderFooterResolvedParagraphText(parts, paragraph, text, pageNumberStyle, forcedZone ?? zoneOverride, listMarkers, nativeFontMap, replacements);
        }

        private static void AddNativeHeaderFooterResolvedParagraphText(
            NativeHeaderFooterText parts,
            WordParagraph paragraph,
            string? text,
            PdfCore.PdfPageNumberStyle? pageNumberStyle,
            NativeHeaderFooterZone? resolvedZone,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            NativeFontMap? nativeFontMap,
            IReadOnlyList<NativeHeaderFooterStyledReplacement>? replacements = null) {
            string resolvedText = text ?? string.Empty;
            PdfCore.PdfTextRun? markerRun = null;
            PdfCore.PdfTextRun? contentStyleRun = null;
            WordDocumentTraversal.ListInfo? listInfo = WordDocumentTraversal.GetListInfo(paragraph);
            if (listInfo != null && listMarkers.TryGetValue(paragraph, out var marker)) {
                string serializedPrefix = marker.Marker + ResolveNativeInlineListMarkerSuffix(listInfo.Value.LevelSuffix);
                if (serializedPrefix.Length > 0 && resolvedText.StartsWith(serializedPrefix, StringComparison.Ordinal)) {
                    resolvedText = resolvedText.Substring(serializedPrefix.Length);
                }

                NativeResolvedTextStyle textStyle = ResolveNativeTextRunStyle(paragraph, nativeFontMap: nativeFontMap);
                (double MarkerOffset, double TextOffset) offsets = ResolveNativeHeaderFooterListOffsets(paragraph, listInfo.Value);
                if (!string.IsNullOrEmpty(marker.Marker)) {
                    markerRun = CreateNativeHeaderFooterListMarkerTextRun(
                        marker.Marker,
                        paragraph,
                        listInfo.Value,
                        textStyle,
                        nativeFontMap,
                        offsets.MarkerOffset,
                        offsets.TextOffset);
                } else if (!string.IsNullOrEmpty(resolvedText) && offsets.TextOffset > 0D) {
                    contentStyleRun = CreateNativeHeaderFooterStyledTextRun(resolvedText, textStyle, offsets.TextOffset);
                }
            }

            if (string.IsNullOrWhiteSpace(resolvedText) && markerRun == null) {
                if (resolvedZone.HasValue) {
                    parts.Append(resolvedZone.Value, string.Empty, pageNumberStyle, markerRun: null, contentStyleRun: null);
                }
                return;
            }

            NativeResolvedTextStyle paragraphTextStyle = ResolveNativeTextRunStyle(paragraph, nativeFontMap: nativeFontMap);
            if (contentStyleRun == null && HasNativeTextSpacing(paragraphTextStyle.TextSpacing)) {
                contentStyleRun = CreateNativeHeaderFooterStyledTextRun(resolvedText, paragraphTextStyle, 0D);
            }

            if (resolvedZone.HasValue) {
                parts.Append(resolvedZone.Value, resolvedText, pageNumberStyle, markerRun, contentStyleRun, replacements);
                return;
            }

            W.JustificationValues? alignment = ResolveNativeParagraphJustification(paragraph);
            if (alignment == W.JustificationValues.Center) {
                parts.AppendCenter(resolvedText, pageNumberStyle, markerRun, contentStyleRun, replacements);
            } else if (alignment == W.JustificationValues.Right) {
                parts.AppendRight(resolvedText, pageNumberStyle, markerRun, contentStyleRun, replacements);
            } else {
                parts.AppendLeft(resolvedText, pageNumberStyle, markerRun, contentStyleRun, replacements);
            }
        }

        private static IReadOnlyList<NativeHeaderFooterStyledReplacement> CreateNativeHeaderFooterTextBoxListReplacements(
            WordTextBox textBox,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            NativeFontMap? nativeFontMap) {
            var replacements = new List<NativeHeaderFooterStyledReplacement>();
            var pending = new Stack<IEnumerator<WordParagraph>>();
            pending.Push(GetNativeTextBoxParagraphs(textBox).GetEnumerator());
            while (pending.Count > 0) {
                IEnumerator<WordParagraph> current = pending.Peek();
                if (!current.MoveNext()) {
                    current.Dispose();
                    pending.Pop();
                    continue;
                }

                WordParagraph innerParagraph = current.Current;
                WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(innerParagraph);
                if (info != null &&
                    listMarkers.TryGetValue(innerParagraph, out var marker) &&
                    !string.IsNullOrEmpty(marker.Marker)) {
                    string serializedPrefix = NormalizeNativeDirectText(
                        marker.Marker + ResolveNativeInlineListMarkerSuffix(info.Value.LevelSuffix));
                    NativeResolvedTextStyle textStyle = ResolveNativeTextRunStyle(innerParagraph, nativeFontMap: nativeFontMap);
                    (double MarkerOffset, double TextOffset) offsets = ResolveNativeHeaderFooterListOffsets(innerParagraph, info.Value);
                    PdfCore.PdfTextRun styledMarker = CreateNativeHeaderFooterListMarkerTextRun(
                        marker.Marker,
                        innerParagraph,
                        info.Value,
                        textStyle,
                        nativeFontMap,
                        offsets.MarkerOffset,
                        offsets.TextOffset);
                    replacements.Add(new NativeHeaderFooterStyledReplacement(serializedPrefix, styledMarker));
                }

                WordTextBox? nested = GetNativeParagraphTextBox(innerParagraph, out _);
                if (nested != null) {
                    pending.Push(GetNativeTextBoxParagraphs(nested).GetEnumerator());
                }
            }
            return replacements;
        }

        private static (double MarkerOffset, double TextOffset) ResolveNativeHeaderFooterListOffsets(
            WordParagraph paragraph,
            WordDocumentTraversal.ListInfo info) {
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            bool useParagraphStyleIndent = ShouldApplyNativeListParagraphStyleIndent(paragraph);
            double textOffset = paragraph.IndentationBeforePoints ??
                (useParagraphStyleIndent ? styleDefaults.LeftIndent : null) ??
                ConvertNativeTwipsToPoints(info.LeftIndentTwips ?? ((info.Level + 1) * 720)) ?? 0D;
            double hangingIndent = paragraph.IndentationHangingPoints ??
                (useParagraphStyleIndent ? GetNativeStyleHangingIndent(styleDefaults) : null) ??
                ConvertNativeTwipsToPoints(info.HangingIndentTwips ?? 360) ?? 0D;
            double markerOffset = Math.Max(0D, textOffset - Math.Max(0D, hangingIndent));
            return (
                Math.Min(MaximumNativeHeaderFooterListOffsetPoints, markerOffset),
                Math.Min(MaximumNativeHeaderFooterListOffsetPoints, Math.Max(0D, textOffset)));
        }

        private static PdfCore.PdfTextRun CreateNativeHeaderFooterListMarkerTextRun(
            string marker,
            WordParagraph paragraph,
            WordDocumentTraversal.ListInfo info,
            NativeResolvedTextStyle textStyle,
            NativeFontMap? nativeFontMap,
            double markerOffset,
            double textOffset) {
            PdfCore.PdfTextRun styledMarker = CreateNativeListMarkerTextRun(marker, paragraph, textStyle, nativeFontMap);
            double markerFontSize = styledMarker.FontSize ?? textStyle.FontSize ?? 12D;
            double markerWidth = EstimateNativeListMarkerWidth(marker, markerFontSize, textStyle.TextSpacing);
            double markerColumnWidth = Math.Max(markerWidth, Math.Max(0D, textOffset - markerOffset));
            double alignmentOffset = info.LevelJustification switch {
                WordListLevelAlignment.Right => Math.Max(0D, markerColumnWidth - markerWidth),
                WordListLevelAlignment.Center => Math.Max(0D, (markerColumnWidth - markerWidth) / 2D),
                _ => 0D
            };
            double resolvedMarkerOffset = Math.Min(
                MaximumNativeHeaderFooterListOffsetPoints,
                markerOffset + alignmentOffset);
            string suffix = ResolveNativeHeaderFooterListMarkerSuffix(
                info.LevelSuffix,
                info.LevelJustification,
                marker,
                markerFontSize,
                resolvedMarkerOffset,
                textOffset,
                textStyle.TextSpacing);
            return CloneNativeHeaderFooterTextRun(styledMarker, marker + suffix)
                .WithHorizontalOffset(resolvedMarkerOffset);
        }

        private static string ResolveNativeHeaderFooterListMarkerSuffix(
            WordListLevelSuffix? suffix,
            WordListLevelAlignment? justification,
            string marker,
            double markerFontSize,
            double markerOffset,
            double textOffset, NativeTextSpacing textSpacing = default) {
            bool leftJustified = justification != WordListLevelAlignment.Right &&
                                 justification != WordListLevelAlignment.Center;
            if (leftJustified && suffix == WordListLevelSuffix.Nothing) {
                return string.Empty;
            }
            if (leftJustified && suffix == WordListLevelSuffix.Space) {
                return " ";
            }

            double spaceWidth = Math.Max(0.01D, EstimateNativeListMarkerWidth(" ", markerFontSize, textSpacing));
            double desiredGap = Math.Max(0D, textOffset - markerOffset - EstimateNativeListMarkerWidth(marker, markerFontSize, textSpacing));
            if (suffix == WordListLevelSuffix.Space) {
                desiredGap += spaceWidth;
            }
            if (desiredGap <= spaceWidth * 0.001D) {
                return string.Empty;
            }
            int spaceCount = Math.Max(1, (int)Math.Ceiling(desiredGap / spaceWidth));
            return new string(' ', Math.Min(MaximumNativeHeaderFooterListSpacingCharacters, spaceCount));
        }

        private static PdfCore.PdfTextRun CreateNativeHeaderFooterStyledTextRun(
            string text,
            NativeResolvedTextStyle style,
            double horizontalOffset) =>
            style.TextSpacing.ApplyTo(new PdfCore.PdfTextRun(
                text.Replace('\t', ' '),
                bold: style.Bold,
                underline: style.Underline,
                color: style.Color,
                italic: style.Italic,
                strike: style.Strike,
                fontSize: style.FontSize,
                font: style.Font,
                baseline: style.Baseline,
                backgroundColor: style.BackgroundColor,
                fontFamily: style.FontFamily,
                underlineStyle: style.UnderlineStyle,
                strikeStyle: style.StrikeStyle)
            .WithHorizontalOffset(horizontalOffset));

        private static PdfCore.PdfTextRun CloneNativeHeaderFooterTextRun(PdfCore.PdfTextRun source, string text) =>
            CopyNativeTextSpacing(source, new PdfCore.PdfTextRun(
                text.Replace('\t', ' '),
                source.Bold,
                source.Underline,
                source.Color,
                source.Italic,
                source.Strike,
                source.FontSize,
                source.Font,
                baseline: source.Baseline,
                backgroundColor: source.BackgroundColor,
                fontFamily: source.FontFamily,
                underlineStyle: source.UnderlineStyle,
                strikeStyle: source.StrikeStyle,
                decorationColor: source.DecorationColor)
            .WithFeatureSettings(source.FeatureSettings)
            .WithHorizontalOffset(source.HorizontalOffset));

        private static string? GetNativeHeaderFooterParagraphText(
            WordParagraph paragraph,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            out PdfCore.PdfPageNumberStyle? pageNumberStyle,
            int textBoxDepth = 0) {
            return GetNativeHeaderFooterParagraphText(paragraph, listMarkers, out pageNumberStyle, out _, textBoxDepth);
        }

        private static string? GetNativeHeaderFooterParagraphText(
            WordParagraph paragraph,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            out PdfCore.PdfPageNumberStyle? pageNumberStyle,
            out NativeHeaderFooterZone? zoneOverride,
            int textBoxDepth = 0) {
            zoneOverride = null;
            WordTextBox? numberedTextBox = GetNativeParagraphTextBox(paragraph, out _);
            if (numberedTextBox != null && GetNativeTextBoxParagraphs(numberedTextBox).Any(inner => inner.IsListItem)) {
                zoneOverride = MapNativeTextBoxHeaderFooterZone(numberedTextBox.HorizontalAlignment);
            }
            if (TryBuildNativeHeaderFooterParagraphText(paragraph, out string? mixedText, out pageNumberStyle)) {
                return PrependNativeHeaderFooterListMarker(paragraph, AppendNativeHeaderFooterSupplementalText(mixedText, paragraph), listMarkers, textBoxDepth);
            }

            if (TryGetNativeHeaderFooterFieldToken(paragraph, out string? fieldToken, out pageNumberStyle)) {
                return PrependNativeHeaderFooterListMarker(paragraph, AppendNativeHeaderFooterSupplementalText(fieldToken, paragraph), listMarkers, textBoxDepth);
            }

            pageNumberStyle = null;
            if (paragraph.IsHyperLink && paragraph.Hyperlink != null && !IsNativeHiddenTextRun(paragraph)) {
                string linkedText = string.Concat(GetNativeRuns(paragraph)
                    .Where(run => ReferenceEquals(run._hyperlink, paragraph._hyperlink) && !IsNativeHiddenTextRun(run, paragraph))
                    .Select(run => ApplyNativeTextTransform(run.Text, run, paragraph)));
                return PrependNativeHeaderFooterListMarker(paragraph,
                    AppendNativeHeaderFooterSupplementalText(linkedText, paragraph), listMarkers, textBoxDepth);
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            string? text = runs.Count > 0
                ? string.Concat(runs.Where(run => !IsNativeHiddenTextRun(run, paragraph)).Select(run => ApplyNativeTextTransform(run.Text, run, paragraph)))
                : IsNativeHiddenTextRun(paragraph) || paragraph._paragraph.Descendants<W.FieldChar>().Any() ||
                    !WordComplexFieldRunVisibility.ForParagraph(paragraph._paragraph).IsVisible
                    ? string.Empty : ApplyNativeTextTransform(paragraph.Text, paragraph);
            text = AppendNativeHeaderFooterSupplementalText(text, paragraph);
            if (!string.IsNullOrWhiteSpace(text)) {
                return PrependNativeHeaderFooterListMarker(paragraph, text, listMarkers, textBoxDepth);
            }

            string? textBoxText = GetNativeParagraphTextBoxPlainText(paragraph);
            if (string.IsNullOrWhiteSpace(textBoxText)) {
                return text;
            }

            WordTextBox? textBox = GetNativeParagraphTextBox(paragraph, out _);
            zoneOverride = MapNativeTextBoxHeaderFooterZone(textBox?.HorizontalAlignment ?? WordTextBoxHorizontalAlignment.Center);
            if (textBox != null) {
                EnsureNativeHeaderFooterTextBoxDepth(textBoxDepth);
                string innerText = string.Join("\n", GetNativeTextBoxParagraphs(textBox)
                    .Select(inner => GetNativeHeaderFooterParagraphText(inner, listMarkers, out _, textBoxDepth + 1))
                    .Where(inner => !string.IsNullOrWhiteSpace(inner)));
                if (!string.IsNullOrWhiteSpace(innerText)) {
                    return PrependNativeHeaderFooterListMarker(paragraph, innerText, listMarkers, textBoxDepth);
                }
            }
            return PrependNativeHeaderFooterListMarker(paragraph, textBoxText, listMarkers, textBoxDepth);
        }

        private static string? PrependNativeHeaderFooterListMarker(
            WordParagraph paragraph,
            string? text,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            int textBoxDepth) {
            if (string.IsNullOrWhiteSpace(text)) return text;
            string resolvedText = text!;
            WordTextBox? textBox = GetNativeParagraphTextBox(paragraph, out _);
            if (textBox != null) {
                IReadOnlyList<WordParagraph> innerParagraphs = GetNativeTextBoxParagraphs(textBox);
                if (innerParagraphs.Any(inner => inner.IsListItem)) {
                    EnsureNativeHeaderFooterTextBoxDepth(textBoxDepth);
                    string? flattenedText = GetNativeParagraphTextBoxPlainText(paragraph);
                    int textBoxStart = string.IsNullOrEmpty(flattenedText)
                        ? -1
                        : resolvedText.IndexOf(flattenedText, StringComparison.Ordinal);
                    if (textBoxStart >= 0) {
                        string innerText = string.Join("\n", innerParagraphs
                            .Select(inner => GetNativeHeaderFooterParagraphText(inner, listMarkers, out _, textBoxDepth + 1))
                            .Where(inner => !string.IsNullOrWhiteSpace(inner)));
                        resolvedText = resolvedText.Remove(textBoxStart, flattenedText!.Length).Insert(textBoxStart, innerText);
                    }
                }
            }
            if (!paragraph.IsListItem) return resolvedText;
            return listMarkers.TryGetValue(paragraph, out var marker) && !string.IsNullOrEmpty(marker.Marker)
                ? marker.Marker + ResolveNativeInlineListMarkerSuffix(WordDocumentTraversal.GetListInfo(paragraph)?.LevelSuffix) + resolvedText
                : resolvedText;
        }

        private static void EnsureNativeHeaderFooterTextBoxDepth(int textBoxDepth) {
            if (textBoxDepth >= MaximumNativeHeaderFooterTextBoxNestingDepth) {
                throw new InvalidDataException(
                    $"Header or footer text-box nesting exceeds the supported limit of {MaximumNativeHeaderFooterTextBoxNestingDepth} levels.");
            }
        }

        private static string? AppendNativeHeaderFooterSupplementalText(string? text, WordParagraph paragraph) {
            text = AppendNativeHeaderFooterEquationText(text, paragraph);
            text = AppendNativeHeaderFooterFormControlText(text, paragraph);
            text = AppendNativeHeaderFooterTextPathText(text, paragraph);
            return AppendNativeHeaderFooterRepeatingSectionText(text, paragraph);
        }

        private static string? AppendNativeHeaderFooterEquationText(string? text, WordParagraph paragraph) {
            IReadOnlyList<WordEquationOccurrence> occurrences = WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph);
            if (occurrences.Count > 0) {
                string orderedText = AppendNativeTextWithEquation(text ?? string.Empty, paragraph);
                if (!string.IsNullOrWhiteSpace(orderedText)) {
                    return orderedText;
                }
            }

            string? equationText = GetNativeEquationText(paragraph);
            if (string.IsNullOrWhiteSpace(equationText)) {
                return text;
            }

            var builder = new StringBuilder(text ?? string.Empty);
            string currentText = builder.ToString();
            AppendNativeHeaderFooterSupplementalValue(builder, ref currentText, equationText, skipIfPresent: true);
            return builder.Length == 0 ? text : builder.ToString();
        }

        private static string? AppendNativeHeaderFooterFormControlText(string? text, WordParagraph paragraph) {
            IReadOnlyList<W.SdtRun> checkBoxes = GetNativeCheckBoxControls(paragraph);
            IReadOnlyList<W.SdtRun> formFields = GetNativeFormFieldControls(paragraph);
            if (checkBoxes.Count == 0 && formFields.Count == 0) {
                return text;
            }

            var builder = new StringBuilder(text ?? string.Empty);
            string currentText = builder.ToString();
            foreach (W.SdtRun checkBox in checkBoxes) {
                AppendNativeHeaderFooterSupplementalValue(
                    builder,
                    ref currentText,
                    IsNativeCheckBoxChecked(checkBox) ? "[x]" : "[ ]",
                    skipIfPresent: false);
            }

            foreach (W.SdtRun formField in formFields) {
                string? value;
                if (IsNativeDatePickerControl(formField)) {
                    value = GetNativeDatePickerValue(formField);
                } else {
                    IReadOnlyList<string> options = GetNativeChoiceFieldOptions(formField);
                    value = GetNativeChoiceFieldValue(formField, options);
                }

                AppendNativeHeaderFooterSupplementalValue(builder, ref currentText, value, skipIfPresent: true);
            }

            return builder.Length == 0 ? text : builder.ToString();
        }

        private static string? AppendNativeHeaderFooterTextPathText(string? text, WordParagraph paragraph) {
            IEnumerable<V.TextPath> textPaths = paragraph._paragraph?.Descendants<V.TextPath>() ?? Enumerable.Empty<V.TextPath>();
            var builder = new StringBuilder(text ?? string.Empty);
            string currentText = builder.ToString();
            foreach (V.TextPath textPath in textPaths) {
                if (!IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(textPath, "on")) ||
                    IsNativeHeaderFooterWatermarkTextPath(textPath)) {
                    continue;
                }

                AppendNativeHeaderFooterSupplementalValue(
                    builder,
                    ref currentText,
                    textPath.String?.Value ?? GetNativeOpenXmlAttribute(textPath, "string"),
                    skipIfPresent: true);
            }

            return builder.Length == 0 ? text : builder.ToString();
        }

        private static bool IsNativeHeaderFooterWatermarkTextPath(V.TextPath textPath) {
            V.Shape? shape = textPath.Ancestors<V.Shape>().FirstOrDefault();
            if (shape == null) {
                return false;
            }

            string marker = string.Join(" ",
                shape.Id?.Value,
                GetNativeOpenXmlAttribute(shape, "name"),
                GetNativeOpenXmlAttribute(shape, "title"));
            return marker.IndexOf("watermark", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string? AppendNativeHeaderFooterRepeatingSectionText(string? text, WordParagraph paragraph) {
            IReadOnlyList<W.SdtRun> controls = GetNativeRepeatingSectionControls(paragraph);
            if (controls.Count == 0) {
                return text;
            }

            var builder = new StringBuilder(text ?? string.Empty);
            string currentText = builder.ToString();
            foreach (W.SdtRun control in controls) {
                foreach (string itemText in GetNativeRepeatingSectionItems(control)) {
                    AppendNativeHeaderFooterSupplementalValue(builder, ref currentText, itemText, skipIfPresent: true);
                }
            }

            return builder.Length == 0 ? text : builder.ToString();
        }

        private static void AppendNativeHeaderFooterSupplementalValue(StringBuilder builder, ref string currentText, string? value, bool skipIfPresent) {
            if (string.IsNullOrWhiteSpace(value) ||
                skipIfPresent && currentText.IndexOf(value!, StringComparison.Ordinal) >= 0) {
                return;
            }

            if (builder.Length > 0 && !char.IsWhiteSpace(builder[builder.Length - 1])) {
                builder.Append(' ');
            }

            builder.Append(value);
            currentText = builder.ToString();
        }

        private static string AppendNativeTextWithEquation(string text, WordParagraph paragraph) {
            IReadOnlyList<WordEquationContentSegment> segments = GetNativeVisibleEquationContentSegments(paragraph);
            if (segments.Count == 0) {
                return text;
            }

            string orderedText = string.Concat(segments.Select(GetNativeEquationSegmentText));
            return string.IsNullOrEmpty(orderedText) ? text : orderedText;
        }

        private static IReadOnlyList<WordEquationContentSegment> GetNativeVisibleEquationContentSegments(WordParagraph paragraph) {
            IReadOnlyList<WordEquationOccurrence> occurrences = WordEquation.GetOccurrences(paragraph._document, paragraph._paragraph);
            if (occurrences.Count == 0 || GetNativeParagraphStyleDefaults(paragraph).Hidden == true) {
                return Array.Empty<WordEquationContentSegment>();
            }

            return WordEquation.GetVisibleContentSegments(
                paragraph._paragraph,
                occurrences,
                element => element is not W.Run run ||
                    !IsNativeHiddenTextRun(new WordParagraph(paragraph._document, paragraph._paragraph, run), paragraph));
        }

        private static string? GetNativeEquationText(WordParagraph paragraph) {
            string[] equationTexts = WordEquation
                .GetOccurrences(paragraph._document, paragraph._paragraph)
                .Select(occurrence => occurrence.Equation.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToArray();
            if (equationTexts.Length > 0) {
                return string.Join(" ", equationTexts);
            }

            string ommlText = WordMath.GetText(paragraph._paragraph);
            return string.IsNullOrWhiteSpace(ommlText) ? null : ommlText;
        }

        private static NativeHeaderFooterZone MapNativeTextBoxHeaderFooterZone(WordTextBoxHorizontalAlignment alignment) {
            switch (alignment) {
                case WordTextBoxHorizontalAlignment.Center:
                    return NativeHeaderFooterZone.Center;
                case WordTextBoxHorizontalAlignment.Right:
                case WordTextBoxHorizontalAlignment.Outside:
                    return NativeHeaderFooterZone.Right;
                default:
                    return NativeHeaderFooterZone.Left;
            }
        }

    }
}
