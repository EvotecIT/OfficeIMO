using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool TryRenderNativeList(
            INativePdfFlow pdf,
            IReadOnlyList<WordElement> elements,
            ref int index,
            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            Dictionary<WordParagraph, (int Level, int Index)> listIndices,
            NativeNoteNumbering footnoteNumbersById,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap,
            IReadOnlyDictionary<W.Paragraph, string> headingDestinations) {
            if (elements[index] is not WordParagraph firstParagraph ||
                !TryGetNativeListItem(firstParagraph, listMarkers, listIndices, footnoteNumbersById, nativeDefaults, nativeFontMap, headingDestinations, out bool ordered, out int level, out int startNumber, out PdfCore.PdfListItem? item, out PdfCore.PdfAlign align, out PdfCore.PdfColor? color, out PdfCore.PdfListStyle? style)) {
                return false;
            }

            var items = new List<PdfCore.PdfListItem> { item! };
            var paragraphs = new List<WordParagraph> { firstParagraph };
            int nextIndex = index + 1;
            int expectedNumber = startNumber + 1;
            // A Word keep-lines constraint applies to one numbered paragraph, not the whole list.
            while (style?.KeepTogether != true && nextIndex < elements.Count &&
                   elements[nextIndex] is WordParagraph paragraph &&
                   TryGetNativeListItem(paragraph, listMarkers, listIndices, footnoteNumbersById, nativeDefaults, nativeFontMap, headingDestinations, out bool nextOrdered, out int nextLevel, out int nextNumber, out PdfCore.PdfListItem? nextItem, out PdfCore.PdfAlign nextAlign, out PdfCore.PdfColor? nextColor, out PdfCore.PdfListStyle? nextStyle) &&
                   nextOrdered == ordered &&
                   nextLevel == level &&
                   nextAlign == align &&
                   nextColor.Equals(color) &&
                   NativeListStylesEquivalent(nextStyle, style) &&
                   NativeListContextualSpacingEquivalent(firstParagraph, paragraph) &&
                   (!ordered || nextNumber == expectedNumber)) {
                items.Add(nextItem!);
                paragraphs.Add(paragraph);
                nextIndex++;
                expectedNumber++;
            }

            style = ApplyNativeListContextualItemSpacing(style, paragraphs);
            bool suppressBefore = ShouldSuppressNativeContextualSpacingBefore(firstParagraph);
            bool suppressAfter = ShouldSuppressNativeContextualSpacingAfter(paragraphs[paragraphs.Count - 1],
                nextIndex < elements.Count ? elements[nextIndex] as WordParagraph : null);
            double? declaredSpacingAfter = style?.SpacingAfter;
            if (style != null && (suppressBefore || suppressAfter)) {
                style = style.Clone();
                if (suppressBefore) style.SpacingBefore = 0D;
                if (suppressAfter) style.SpacingAfter = 0D;
            }
            if (style != null && RenderNativeDocumentStartSpacing(pdf, firstParagraph, style.SpacingBefore)) {
                style = style.Clone();
                style.SpacingBefore = 0D;
            }
            if (ordered) {
                pdf.RichNumbered(items, align, color, startNumber, style);
            } else {
                pdf.RichBullets(items, align, color, style);
            }
            if (suppressAfter && pdf is NativeSpacingCollapseFlow collapseFlow)
                collapseFlow.RetainSuppressedSpacingAfter(declaredSpacingAfter);

            index = nextIndex - 1;
            return true;
        }

        private static bool TryGetNativeListItem(
            WordParagraph paragraph,
            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            Dictionary<WordParagraph, (int Level, int Index)> listIndices,
            NativeNoteNumbering footnoteNumbersById,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap,
            IReadOnlyDictionary<W.Paragraph, string> headingDestinations,
            out bool ordered,
            out int level,
            out int index,
            out PdfCore.PdfListItem? item,
            out PdfCore.PdfAlign align,
            out PdfCore.PdfColor? color,
            out PdfCore.PdfListStyle? style) {
            ordered = false;
            level = 0;
            index = 1;
            item = null;
            align = PdfCore.PdfAlign.Left;
            color = null;
            style = null;

            if (!listMarkers.TryGetValue(paragraph, out var marker) ||
                !listIndices.TryGetValue(paragraph, out var listIndex)) {
                return false;
            }

            WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraph);
            if (info == null || !info.Value.MarkerVisible || marker.Level != info.Value.Level || listIndex.Level != info.Value.Level) {
                return false;
            }

            // Space and nothing suffixes follow each item's actual marker advance.
            // The paragraph path retains that per-item position instead of a shared column.
            if (info.Value.LevelSuffix is WordListLevelSuffix.Space or WordListLevelSuffix.Nothing) return false;

            // Numbering tabs may place the first line before or after the
            // continuation indent. Paragraphs preserve both positions;
            // list blocks expose one shared text column.
            if (HasNativeParagraphBarTabs(paragraph) || IgnoresNativeNumberingIndentStop(paragraph) ||
                HasNativeIndependentListTabAdvance(paragraph, marker.Marker, nativeDefaults, nativeFontMap)) return false;

            // The inline paragraph path carries marker run typography. List blocks
            // expose a uniform marker font but do not carry width or tracking.
            NativeResolvedTextStyle paragraphTextStyle = ResolveNativeTextRunStyle(paragraph,
                nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            NativeTextSpacing markerSpacing = ResolveNativeListMarkerTextSpacing(info.Value, paragraphTextStyle.ListMarkerTextSpacing);
            if (HasNativeTextSpacing(paragraphTextStyle.TextSpacing) ||
                markerSpacing.WidthPercentage != 100D || markerSpacing.CharacterSpacing != 0D) return false;

            // A differently sized or styled label changes only the first line's
            // metrics. Paragraph runs retain that boundary across wrapped lines.
            if (!DoubleEquals(info.Value.MarkerFontSize ?? paragraphTextStyle.EffectiveFontSize,
                    paragraphTextStyle.EffectiveFontSize) ||
                ResolveNativeListMarkerFont(info.Value, marker.Marker, paragraphTextStyle, nativeFontMap) != paragraphTextStyle.Font ||
                !string.Equals(ResolveNativeListMarkerFontFamily(info.Value, marker.Marker, paragraphTextStyle, nativeFontMap),
                    paragraphTextStyle.FontFamily, StringComparison.OrdinalIgnoreCase) ||
                (info.Value.MarkerBold ?? paragraphTextStyle.Bold) != paragraphTextStyle.Bold ||
                (info.Value.MarkerItalic ?? paragraphTextStyle.Italic) != paragraphTextStyle.Italic) return false;

            if (HasNativePageBreakBefore(paragraph) ||
                paragraph.IsPageBreak ||
                paragraph.Shape != null ||
                paragraph.TextBox != null ||
                paragraph.Chart != null ||
                paragraph.PictureControl?.Image != null ||
                GetNativePictureControls(paragraph).Any(sdtRun => IsNativePictureControlWithImage(paragraph, sdtRun)) ||
                paragraph.Image != null) {
                return false;
            }

            List<WordParagraph> runs = GetNativeRuns(paragraph);
            if (runs.Any(run => !IsNativeHiddenTextRun(run, paragraph) &&
                run.GetNonTextBreakPositions()?.Values.Any(type => type == WordBreakType.Page || type == WordBreakType.Column) == true)) return false;
            if (runs.Any(run => run.IsImage) || HasNativeParagraphShapeGroups(runs)) {
                return false;
            }

            List<PdfCore.PdfTextRun> richRuns = CreateNativeCellParagraphRuns(paragraph, footnoteNumbersById, NativeTableStyleDefaults.Empty, nativeDefaults, nativeFontMap);
            string content = string.Concat(richRuns.Select(run => run.Text));
            if (string.IsNullOrWhiteSpace(content)) {
                return false;
            }

            bool itemOrdered = info.Value.Ordered;
            string displayMarker = itemOrdered
                ? marker.Marker
                : NormalizeNativeBulletMarker(marker.Marker);
            ordered = itemOrdered;
            level = info.Value.Level;
            index = listIndex.Index;
            string? destination = paragraph.Bookmark?.Name;
            if (string.IsNullOrEmpty(destination)) headingDestinations.TryGetValue(paragraph._paragraph, out destination);
            item = new PdfCore.PdfListItem(richRuns, destination, string.IsNullOrWhiteSpace(displayMarker) ? null : displayMarker);
            align = ResolveNativeParagraphAlign(paragraph, allowJustify: false);
            NativeResolvedTextStyle textStyle = ResolveNativeTextRunStyle(paragraph, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            color = textStyle.Color;
            style = CreateNativeListStyle(paragraph, info.Value, displayMarker, nativeDefaults, textStyle, nativeFontMap);
            PdfCore.PdfTextRun measuredMarker = CreateNativeListMarkerTextRun(displayMarker, paragraph,
                textStyle, nativeFontMap, includeSuffix: false);
            double markerWidth = nativeFontMap.MeasureText(measuredMarker)
                ?? EstimateNativeListMarkerWidth(displayMarker, measuredMarker.FontSize ?? nativeDefaults.FontSize);
            // An overrun advances only the first line to the next Word tab stop.
            // The paragraph path keeps continuation lines at the authored indent.
            if (markerWidth - GetNativeMarkerAnchorShift(info.Value, markerWidth) > style.MarkerWidth + 0.01D)
                return false;
            return true;
        }

        private static string NormalizeNativeBulletMarker(string marker) {
            if (string.IsNullOrWhiteSpace(marker)) {
                return "•";
            }

            return marker.Trim() switch {
                "\uf0b7" => "•",
                "\u00b7" => "•",
                "\u25cf" => "•",
                "\u006f" => "o",
                _ => marker
            };
        }

        private static PdfCore.PdfListStyle CreateNativeListStyle(WordParagraph paragraph, WordDocumentTraversal.ListInfo info, string marker, NativeDocumentDefaults nativeDefaults, NativeResolvedTextStyle markerTextStyle, NativeFontMap nativeFontMap) {
            const double defaultLevelTextIndent = 36D;
            const double defaultHangingIndent = 18D;
            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);

            double numberingTextIndent = ConvertNativeTwipsToPoints(info.LeftIndentTwips ?? ((info.Level + 1) * 720)) ??
                ((info.Level + 1) * defaultLevelTextIndent);
            double numberingHangingIndent = ConvertNativeTwipsToPoints(info.HangingIndentTwips ?? 360) ??
                defaultHangingIndent;
            bool useParagraphStyleIndent = ShouldApplyNativeListParagraphStyleIndent(paragraph);
            double textIndent = paragraph.IndentationBeforePoints ??
                (useParagraphStyleIndent ? styleDefaults.LeftIndent : null) ??
                numberingTextIndent;
            double hangingIndent = paragraph.IndentationHangingPoints ??
                (useParagraphStyleIndent ? GetNativeStyleHangingIndent(styleDefaults) : null) ??
                numberingHangingIndent;
            double markerIndent = Math.Max(0D, textIndent - hangingIndent);
            double fontSize = ResolveNativeParagraphEffectiveFontSize(paragraph, nativeDefaults, styleDefaults);
            double naturalLineHeight = ResolveNativeParagraphSingleLineHeight(paragraph, nativeDefaults, styleDefaults, nativeFontMap: nativeFontMap);
            NativeLineSpacing lineSpacing = ResolveNativeParagraphLineSpacing(paragraph, styleDefaults, nativeDefaults);
            double lineHeight = lineSpacing.Resolve(fontSize, naturalLineHeight) ?? nativeDefaults.ParagraphLineHeight;
            W.SpacingBetweenLines? directSpacing = paragraph._paragraph?.ParagraphProperties?.GetFirstChild<W.SpacingBetweenLines>();
            double markerFontSize = info.MarkerFontSize ?? fontSize;
            bool itemSpacingDeclared = false;

            var style = new PdfCore.PdfListStyle {
                LeftIndent = markerIndent,
                MarkerAlignsAtIndent = true,
                MarkerFeatureSettings = default(NativeTextSpacing).ResolveFeatureSettings(markerFontSize),
                MarkerGap = 0D,
                MarkerWidth = Math.Max(0D, textIndent - markerIndent),
                MarkerFont = ResolveNativeListMarkerFont(info, marker, markerTextStyle, nativeFontMap),
                MarkerFontFamily = ResolveNativeListMarkerFontFamily(info, marker, markerTextStyle, nativeFontMap),
                MarkerFontSize = info.MarkerFontSize ?? ResolveNativeParagraphFontSize(paragraph, nativeDefaults, styleDefaults),
                MarkerColor = ParseNativeColor(info.MarkerColorHex),
                MarkerAlign = MapNativeListMarkerAlign(info.LevelJustification) ?? PdfCore.PdfAlign.Left,
                MarkerBold = info.MarkerBold ?? markerTextStyle.Bold,
                MarkerItalic = info.MarkerItalic ?? markerTextStyle.Italic
            };

            style.LineHeight = lineHeight;
            style.LineSpacing = lineSpacing.ToPdfLineSpacing(naturalLineHeight);
            style.FontSize = ResolveNativeParagraphLayoutFontSize(paragraph, nativeDefaults, styleDefaults);

            var spacingInLines = ResolveNativeParagraphSpacingInLines(paragraph, styleDefaults, nativeDefaults);
            if (spacingInLines.Before.HasValue) {
                style.SpacingBefore = spacingInLines.Before.Value;
            } else if (paragraph.LineSpacingBeforePoints.HasValue) {
                style.SpacingBefore = paragraph.LineSpacingBeforePoints.Value;
            } else if (GetNativeSpacingBeforePoints(directSpacing, fontSize, lineHeight) is { } directSpacingBefore) {
                style.SpacingBefore = directSpacingBefore;
            } else if (styleDefaults.SpacingBefore.HasValue) {
                style.SpacingBefore = styleDefaults.SpacingBefore.Value;
            } else if (nativeDefaults.ParagraphSpacingBeforeDeclared) {
                style.SpacingBefore = nativeDefaults.ParagraphSpacingBefore;
            }

            if (spacingInLines.After.HasValue) {
                style.SpacingAfter = spacingInLines.After.Value;
                itemSpacingDeclared = true;
            } else if (paragraph.LineSpacingAfterPoints.HasValue) {
                style.SpacingAfter = paragraph.LineSpacingAfterPoints.Value;
                itemSpacingDeclared = true;
            } else if (GetNativeSpacingAfterPoints(directSpacing, fontSize, lineHeight) is { } directSpacingAfter) {
                style.SpacingAfter = directSpacingAfter;
                itemSpacingDeclared = true;
            } else if (styleDefaults.SpacingAfter.HasValue) {
                style.SpacingAfter = styleDefaults.SpacingAfter.Value;
                itemSpacingDeclared = true;
            } else if (nativeDefaults.ParagraphSpacingAfterDeclared) {
                style.SpacingAfter = nativeDefaults.ParagraphSpacingAfter;
                itemSpacingDeclared = true;
            } else {
                style.SpacingAfter = nativeDefaults.ParagraphSpacingAfter;
            }

            if (itemSpacingDeclared || style.SpacingBefore > 0D) {
                // Each item is a Word paragraph. Its before/after spacing collapses
                // to the larger gap even when several paragraphs share a PDF list.
                style.ItemSpacing = Math.Max(style.SpacingBefore, style.SpacingAfter ?? 0D);
            }

            style.KeepTogether = ReadNativeDirectParagraphOnOff<W.KeepLines>(paragraph) ?? styleDefaults.KeepTogether ?? false;
            style.KeepTogetherOverflow = UsesModernNativeWordLayout(paragraph._document)
                ? PdfCore.PdfKeepTogetherOverflowBehavior.SplitAcrossPages : PdfCore.PdfKeepTogetherOverflowBehavior.SplitAcrossFrames;
            style.KeepWithNext = ReadNativeDirectParagraphOnOff<W.KeepNext>(paragraph) ?? styleDefaults.KeepWithNext ?? false;
            return style;
        }

        private static PdfCore.PdfListStyle? ApplyNativeListContextualItemSpacing(PdfCore.PdfListStyle? style, IReadOnlyList<WordParagraph> paragraphs) {
            if (style == null || paragraphs.Count < 2) {
                return style;
            }

            for (int i = 0; i < paragraphs.Count - 1; i++) {
                if (!ShouldSuppressNativeContextualSpacingAfter(paragraphs[i], paragraphs[i + 1])) {
                    return style;
                }
            }

            PdfCore.PdfListStyle contextualStyle = style.Clone();
            contextualStyle.ItemSpacing = 0D;
            return contextualStyle;
        }

        private static bool NativeListContextualSpacingEquivalent(WordParagraph first, WordParagraph next) {
            bool contextual = UsesNativeContextualSpacing(first);
            return contextual == UsesNativeContextualSpacing(next) && (!contextual ||
                string.Equals(GetNativeParagraphStyleIdentity(first), GetNativeParagraphStyleIdentity(next), StringComparison.Ordinal));
        }

        private static double? GetNativeStyleHangingIndent(NativeParagraphStyleDefaults styleDefaults) =>
            styleDefaults.FirstLineIndent is < 0D ? -styleDefaults.FirstLineIndent.Value : null;

        private static bool ShouldApplyNativeListParagraphStyleIndent(WordParagraph paragraph) =>
            !string.IsNullOrWhiteSpace(paragraph.StyleId) &&
            !string.Equals(paragraph.StyleId, "ListParagraph", StringComparison.OrdinalIgnoreCase);

        private static void ApplyNativeMarkerlessListIndent(WordParagraph paragraph, PdfCore.PdfParagraphStyle style) {
            WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraph);
            if (info == null || info.Value.MarkerVisible) return;
            ApplyNativeInlineListIndent(paragraph, style);
        }

        private static void ApplyNativeInlineListIndent(WordParagraph paragraph, PdfCore.PdfParagraphStyle style) {
            WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraph);
            if (info == null) return;

            NativeParagraphStyleDefaults styleDefaults = GetNativeParagraphStyleDefaults(paragraph);
            bool useParagraphStyleIndent = ShouldApplyNativeListParagraphStyleIndent(paragraph);
            style.LeftIndent = paragraph.IndentationBeforePoints ??
                (useParagraphStyleIndent ? styleDefaults.LeftIndent : null) ??
                ConvertNativeTwipsToPoints(info.Value.LeftIndentTwips ?? ((info.Value.Level + 1) * 720)) ?? 0D;
            double hangingIndent = paragraph.IndentationHangingPoints ??
                (useParagraphStyleIndent ? GetNativeStyleHangingIndent(styleDefaults) : null) ??
                ConvertNativeTwipsToPoints(info.Value.HangingIndentTwips ?? 360) ?? 0D;
            style.FirstLineIndent = paragraph.IndentationHangingPoints.HasValue
                ? -hangingIndent
                : paragraph.IndentationFirstLinePoints ??
                    (useParagraphStyleIndent ? styleDefaults.FirstLineIndent : null) ?? -hangingIndent;
            // Rebase against the resolved continuation frame, after numbering
            // indentation has replaced the paragraph style's provisional value.
            RefreshNativeParagraphTabStops(paragraph, style);
        }

        private static string ResolveNativeInlineListMarkerSuffix(WordListLevelSuffix? suffix) =>
            WordDocumentTraversal.ResolveTextListMarkerSuffix(suffix);

        private static double GetNativeMarkerAnchorShift(WordDocumentTraversal.ListInfo info, double width) =>
            info.LevelJustification switch {
                WordListLevelAlignment.Right => width,
                WordListLevelAlignment.Center => width / 2D,
                _ => 0D
            };

        private static void ApplyNativeInlineListMarkerAlignment(WordParagraph paragraph, string marker,
            PdfCore.PdfParagraphStyle style, NativeDocumentDefaults nativeDefaults, NativeFontMap nativeFontMap) {
            WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraph);
            if (info == null || marker.Length == 0) return;
            NativeResolvedTextStyle textStyle = ResolveNativeTextRunStyle(paragraph,
                nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            PdfCore.PdfTextRun markerRun = CreateNativeListMarkerTextRun(marker, paragraph, textStyle,
                nativeFontMap, includeSuffix: false);
            if (nativeFontMap.MeasureText(markerRun) is { } width) {
                if (info.Value.LevelSuffix == WordListLevelSuffix.Space)
                    width += ResolveNativeListSpaceSuffixWidth(paragraph, nativeDefaults, nativeFontMap);
                style.FirstLineIndent -= GetNativeMarkerAnchorShift(info.Value, width);
            }
        }

        private static PdfCore.PdfTextRun CreateNativeListMarkerTextRun(
            string marker,
            WordParagraph paragraph,
            NativeResolvedTextStyle textStyle,
            NativeFontMap? nativeFontMap,
            bool includeSuffix = true) {
            WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraph);
            if (info == null) {
                return textStyle.ApplyTo(new PdfCore.PdfTextRun(marker + (includeSuffix ? " " : string.Empty), bold: textStyle.Bold, color: textStyle.Color,
                    italic: textStyle.Italic, fontSize: textStyle.FontSize, font: textStyle.Font,
                    fontFamily: textStyle.FontFamily));
            }

            return ResolveNativeListMarkerTextSpacing(info.Value, textStyle.ListMarkerTextSpacing).ApplyTo(new PdfCore.PdfTextRun(
                marker + (includeSuffix ? ResolveNativeInlineListMarkerSuffix(info.Value.LevelSuffix) : string.Empty),
                bold: info.Value.MarkerBold ?? textStyle.Bold,
                color: ParseNativeColor(info.Value.MarkerColorHex) ?? textStyle.Color,
                italic: info.Value.MarkerItalic ?? textStyle.Italic,
                strike: textStyle.Strike,
                fontSize: info.Value.MarkerFontSize ?? textStyle.FontSize,
                font: ResolveNativeListMarkerFont(info.Value, marker, textStyle, nativeFontMap),
                fontFamily: nativeFontMap != null
                    ? ResolveNativeListMarkerFontFamily(info.Value, marker, textStyle, nativeFontMap)
                    : textStyle.FontFamily, strikeStyle: textStyle.StrikeStyle) { ContributesLineBoxDescent = false }, textStyle.EffectiveFontSize);
        }

        private static double EstimateNativeListMarkerWidth(string marker, double fontSize, NativeTextSpacing textSpacing = default) {
            if (string.IsNullOrEmpty(marker)) {
                return 0D;
            }

            double width = 0D;
            foreach (char ch in marker) {
                if (char.IsDigit(ch) || char.IsLetter(ch)) {
                    width += fontSize * 0.56D;
                } else if (char.IsWhiteSpace(ch)) {
                    width += fontSize * 0.28D;
                } else if (ch == '.' || ch == ')' || ch == '(') {
                    width += fontSize * 0.28D;
                } else if (ch == '\u2022' || ch == '\u25CF' || ch == '\u25E6') {
                    width += fontSize * 0.36D;
                } else {
                    width += fontSize * 0.5D;
                }
            }

            return width * (textSpacing.WidthPercentage ?? 100D) / 100D
                + marker.Length * (textSpacing.CharacterSpacing ?? 0D);
        }

        private static bool NativeListStylesEquivalent(PdfCore.PdfListStyle? left, PdfCore.PdfListStyle? right) {
            if (ReferenceEquals(left, right)) {
                return true;
            }

            if (left == null || right == null) {
                return false;
            }

            return NullableDoubleEquals(left.FontSize, right.FontSize) &&
                   NullableDoubleEquals(left.LineHeight, right.LineHeight) &&
                   NativeLineSpacingsEquivalent(left.LineSpacing, right.LineSpacing) &&
                   DoubleEquals(left.LeftIndent, right.LeftIndent) &&
                   NullableDoubleEquals(left.MarkerGap, right.MarkerGap) &&
                   NullableDoubleEquals(left.MarkerWidth, right.MarkerWidth) &&
                   left.MarkerAlignsAtIndent == right.MarkerAlignsAtIndent &&
                   Equals(left.MarkerFeatureSettings, right.MarkerFeatureSettings) &&
                   DoubleEquals(left.SpacingBefore, right.SpacingBefore) &&
                   NullableDoubleEquals(left.SpacingAfter, right.SpacingAfter) &&
                   NullableDoubleEquals(left.ItemSpacing, right.ItemSpacing) &&
                   left.MarkerColor.Equals(right.MarkerColor) &&
                   left.MarkerAlign == right.MarkerAlign &&
                   left.MarkerFont == right.MarkerFont &&
                   string.Equals(left.MarkerFontFamily, right.MarkerFontFamily, StringComparison.OrdinalIgnoreCase) &&
                   NullableDoubleEquals(left.MarkerFontSize, right.MarkerFontSize) &&
                   left.MarkerBold == right.MarkerBold &&
                   left.MarkerItalic == right.MarkerItalic &&
                   left.Color.Equals(right.Color) &&
                   left.KeepTogether == right.KeepTogether &&
                   left.KeepTogetherOverflow == right.KeepTogetherOverflow &&
                   left.KeepWithNext == right.KeepWithNext;
        }

        private static bool NativeLineSpacingsEquivalent(PdfCore.PdfLineSpacing? left, PdfCore.PdfLineSpacing? right) =>
            ReferenceEquals(left, right) || left != null && right != null &&
            left.Rule == right.Rule && DoubleEquals(left.Value, right.Value) &&
            DoubleEquals(left.NaturalMultiplier, right.NaturalMultiplier) &&
            NullableDoubleEquals(left.FontLineBoxMultiplier, right.FontLineBoxMultiplier) &&
            NullableDoubleEquals(left.LineGridPitch, right.LineGridPitch) &&
            left.FontAlignment == right.FontAlignment &&
            DoubleEquals(left.FontAlignmentPointsPerPixel, right.FontAlignmentPointsPerPixel) &&
            NullableDoubleEquals(left.FixedLineBoxBaselineOffset, right.FixedLineBoxBaselineOffset);

        private static bool NullableDoubleEquals(double? left, double? right) {
            if (left.HasValue != right.HasValue) {
                return false;
            }

            return !left.HasValue || DoubleEquals(left.Value, right!.Value);
        }

        private static bool DoubleEquals(double left, double right) =>
            Math.Abs(left - right) < 0.001D;

        private static PdfCore.PdfStandardFont? ResolveNativeListMarkerFont(WordDocumentTraversal.ListInfo info, string marker, NativeResolvedTextStyle markerTextStyle, NativeFontMap? nativeFontMap) {
            if (ShouldUseNativeListTextFontForNormalizedMarker(info, marker)) {
                return markerTextStyle.Font;
            }

            return TryResolveNativeMappedFont(info.MarkerFontFamily, nativeFontMap, out PdfCore.PdfStandardFont markerFont)
                ? markerFont
                : markerTextStyle.Font;
        }

        private static string? ResolveNativeListMarkerFontFamily(WordDocumentTraversal.ListInfo info, string marker, NativeResolvedTextStyle markerTextStyle, NativeFontMap nativeFontMap) {
            if (ShouldUseNativeListTextFontForNormalizedMarker(info, marker)) {
                return markerTextStyle.FontFamily;
            }

            if (nativeFontMap.TryGetNamedFontFamily(info.MarkerFontFamily, out string? markerFamily)) {
                return markerFamily;
            }

            if (nativeFontMap.TryGetFontSlot(info.MarkerFontFamily, out _)) return null;

            return markerTextStyle.FontFamily;
        }

        private static bool ShouldUseNativeListTextFontForNormalizedMarker(WordDocumentTraversal.ListInfo info, string marker) {
            return WordDocumentTraversal.ShouldUseTextFontForMarker(info, marker);
        }

        private static PdfCore.PdfAlign? MapNativeListMarkerAlign(W.LevelJustificationValues? value) {
            if (!value.HasValue) {
                return null;
            }

            if (value.Value == W.LevelJustificationValues.Center) {
                return PdfCore.PdfAlign.Center;
            }

            if (value.Value == W.LevelJustificationValues.Right) {
                return PdfCore.PdfAlign.Right;
            }

            return value.Value == W.LevelJustificationValues.Left ? PdfCore.PdfAlign.Left : null;
        }

        private static PdfCore.PdfAlign? MapNativeListMarkerAlign(WordListLevelAlignment? value) =>
            MapNativeListMarkerAlign(value.ToOpenXml());

        private static List<WordParagraph> GetNativeRuns(WordParagraph paragraph, bool includeInlineContentControls = false) {
            if (paragraph._paragraph == null) {
                return new List<WordParagraph>();
            }

            var runs = new List<WordParagraph>();
            var fieldVisibility = WordComplexFieldRunVisibility.ForParagraph(paragraph._paragraph);
            AddNativeVisibleRuns(runs, paragraph, paragraph._paragraph, null, fieldVisibility, includeInlineContentControls);

            return runs;
        }

        private static void AddNativeVisibleRuns(List<WordParagraph> runs, WordParagraph paragraph,
            DocumentFormat.OpenXml.OpenXmlCompositeElement container, W.Hyperlink? hyperlink,
            WordComplexFieldRunVisibility fieldVisibility, bool includeInlineContentControls = false) {
            const int maxDepth = 128;
            const int maxElements = 1_000_000;
            int visited = 0;
            var pending = new Stack<(IEnumerator<DocumentFormat.OpenXml.OpenXmlElement> Children, W.Hyperlink? Hyperlink, int Depth)>();
            pending.Push((container.ChildElements.GetEnumerator(), hyperlink, 0));
            try {
                while (pending.Count > 0) {
                    var frame = pending.Peek();
                    if (!frame.Children.MoveNext()) {
                        pending.Pop().Children.Dispose();
                        continue;
                    }
                    if (++visited > maxElements)
                        throw new InvalidDataException("Word visible run traversal exceeds the PDF export limit.");
                    var element = frame.Children.Current;
                    DocumentFormat.OpenXml.OpenXmlCompositeElement? nested = null;
                    W.Hyperlink? nestedHyperlink = frame.Hyperlink;
                    if (element is W.Run sourceRun) {
                        W.Run? visibleRun = fieldVisibility.GetVisibleRun(sourceRun, out var visibleSourceChildren);
                        if (visibleRun != null)
                            AppendNativeVisibleRunContent(runs, new WordParagraph(paragraph._document, paragraph._paragraph!, sourceRun) {
                                _hyperlink = frame.Hyperlink,
                                _visibleRun = ReferenceEquals(visibleRun, sourceRun) ? null : visibleRun,
                                _visibleRunSourceChildren = visibleSourceChildren
                            });
                    } else if (element is W.Hyperlink link) {
                        nested = link;
                        nestedHyperlink = link;
                    } else if (element is W.SimpleField simpleField) {
                        if (fieldVisibility.IsVisible)
                            nested = simpleField;
                        else
                            fieldVisibility.ObserveDescendantRuns(simpleField);
                    } else if (element is W.SdtRun sdtRun) {
                        if (includeInlineContentControls && CanProjectNativeInlineContentControl(sdtRun)) {
                            if (fieldVisibility.IsVisible) {
                                if (GetNativeSdtText(sdtRun) == null && !IsNativeCheckBoxControl(sdtRun))
                                    runs.Add(CreateNativeInlineControlFallbackRun(paragraph, sdtRun, frame.Hyperlink));
                                else nested = sdtRun.SdtContentRun;
                            } else fieldVisibility.ObserveDescendantRuns(sdtRun);
                        } else if (!IsNativeSimpleTextContentControl(sdtRun)) {
                            fieldVisibility.ObserveDescendantRuns(sdtRun);
                        } else if (sdtRun.Descendants<W.FieldChar>().Any()) {
                            nested = sdtRun.SdtContentRun;
                        } else if (fieldVisibility.IsVisible) {
                            if (TryGetNativeSdtRunPropertyValue(paragraph._document, sdtRun, out string? propertyValue)) {
                                W.Run resolvedRun = CreateNativeResolvedSdtRun(sdtRun, propertyValue!);
                                runs.Add(new WordParagraph(paragraph._document, paragraph._paragraph!, resolvedRun) { _hyperlink = frame.Hyperlink });
                            } else {
                                nested = sdtRun.SdtContentRun;
                            }
                        }
                    } else if (element is W.CustomXmlRun customXml) {
                        nested = customXml;
                    }
                    if (nested == null) continue;
                    if (frame.Depth >= maxDepth)
                        throw new InvalidDataException("Word visible run nesting exceeds the PDF export limit.");
                    pending.Push((nested.ChildElements.GetEnumerator(), nestedHyperlink, frame.Depth + 1));
                }
            } finally {
                while (pending.Count > 0) pending.Pop().Children.Dispose();
            }
        }

        private static bool TryGetNativeSdtRunPropertyValue(WordDocument document, W.SdtRun sdtRun, out string? value) {
            if (sdtRun.SdtProperties == null || !IsNativePropertyBoundStructuredBlock(sdtRun.SdtProperties)) {
                value = null;
                return false;
            }

            value = GetNativeBuiltInPropertyValue(document, sdtRun.SdtProperties);
            return !string.IsNullOrWhiteSpace(value);
        }

        private static W.Run CreateNativeResolvedSdtRun(W.SdtRun sdtRun, string value) {
            W.Run? sourceRun = sdtRun.SdtContentRun?.Elements<W.Run>().FirstOrDefault();
            var resolvedRun = new W.Run();
            if (sourceRun?.RunProperties != null) {
                resolvedRun.Append((W.RunProperties)sourceRun.RunProperties.CloneNode(true));
            }

            resolvedRun.Append(new W.Text(value) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
            return resolvedRun;
        }

    }
}
