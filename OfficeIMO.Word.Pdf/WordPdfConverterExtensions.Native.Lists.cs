using System.Collections.Generic;
using System.Globalization;
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
            Dictionary<long, int> footnoteNumbersById,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap) {
            if (elements[index] is not WordParagraph firstParagraph ||
                !TryGetNativeListItem(firstParagraph, listMarkers, listIndices, footnoteNumbersById, nativeDefaults, nativeFontMap, out bool ordered, out int level, out int startNumber, out PdfCore.PdfListItem? item, out PdfCore.PdfAlign align, out PdfCore.PdfColor? color, out PdfCore.PdfListStyle? style)) {
                return false;
            }

            var items = new List<PdfCore.PdfListItem> { item! };
            var paragraphs = new List<WordParagraph> { firstParagraph };
            int nextIndex = index + 1;
            int expectedNumber = startNumber + 1;
            while (nextIndex < elements.Count &&
                   elements[nextIndex] is WordParagraph paragraph &&
                   TryGetNativeListItem(paragraph, listMarkers, listIndices, footnoteNumbersById, nativeDefaults, nativeFontMap, out bool nextOrdered, out int nextLevel, out int nextNumber, out PdfCore.PdfListItem? nextItem, out PdfCore.PdfAlign nextAlign, out PdfCore.PdfColor? nextColor, out PdfCore.PdfListStyle? nextStyle) &&
                   nextOrdered == ordered &&
                   nextLevel == level &&
                   nextAlign == align &&
                   nextColor.Equals(color) &&
                   NativeListStylesEquivalent(nextStyle, style) &&
                   (!ordered || nextNumber == expectedNumber)) {
                items.Add(nextItem!);
                paragraphs.Add(paragraph);
                nextIndex++;
                expectedNumber++;
            }

            style = ApplyNativeListContextualItemSpacing(style, paragraphs);
            if (ordered) {
                pdf.RichNumbered(items, align, color, startNumber, style);
            } else {
                pdf.RichBullets(items, align, color, style);
            }

            index = nextIndex - 1;
            return true;
        }

        private static bool TryGetNativeListItem(
            WordParagraph paragraph,
            Dictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            Dictionary<WordParagraph, (int Level, int Index)> listIndices,
            Dictionary<long, int> footnoteNumbersById,
            NativeDocumentDefaults nativeDefaults,
            NativeFontMap nativeFontMap,
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
            if (runs.Any(run => run.IsImage)) {
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
            item = new PdfCore.PdfListItem(richRuns, paragraph.Bookmark?.Name, string.IsNullOrWhiteSpace(displayMarker) ? null : displayMarker);
            align = ResolveNativeParagraphAlign(paragraph, allowJustify: false);
            NativeResolvedTextStyle textStyle = ResolveNativeTextRunStyle(paragraph, nativeDefaults: nativeDefaults, nativeFontMap: nativeFontMap);
            color = textStyle.Color;
            style = CreateNativeListStyle(paragraph, info.Value, displayMarker, nativeDefaults, textStyle, nativeFontMap);
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
            double lineHeight = ResolveNativeParagraphLineHeight(
                paragraph,
                fontSize,
                nativeDefaults,
                styleDefaults,
                nativeFontMap);
            W.SpacingBetweenLines? directSpacing = paragraph._paragraph?.ParagraphProperties?.GetFirstChild<W.SpacingBetweenLines>();
            double markerFontSize = info.MarkerFontSize ?? fontSize;
            double markerTextWidth = EstimateNativeListMarkerWidth(marker, markerFontSize);
            (double markerWidth, double markerGap) = ResolveNativeListMarkerSpacing(info.LevelSuffix, markerTextWidth, markerFontSize, textIndent, markerIndent);
            bool itemSpacingDeclared = false;

            var style = new PdfCore.PdfListStyle {
                LeftIndent = markerIndent,
                MarkerGap = markerGap,
                MarkerWidth = markerWidth,
                MarkerFont = ResolveNativeListMarkerFont(info, marker, markerTextStyle),
                MarkerFontFamily = ResolveNativeListMarkerFontFamily(info, marker, markerTextStyle, nativeFontMap),
                MarkerFontSize = info.MarkerFontSize,
                MarkerColor = ParseNativeColor(info.MarkerColorHex),
                MarkerAlign = MapNativeListMarkerAlign(info.LevelJustification),
                MarkerBold = info.MarkerBold ?? markerTextStyle.Bold,
                MarkerItalic = info.MarkerItalic ?? markerTextStyle.Italic
            };

            if (paragraph.FontSize.HasValue && paragraph.FontSize.Value > 0D) {
                style.FontSize = paragraph.FontSize.Value;
            } else if (styleDefaults.FontSize.HasValue) {
                style.FontSize = styleDefaults.FontSize.Value;
            }

            style.LineHeight = lineHeight;

            if (paragraph.LineSpacingBeforePoints.HasValue) {
                style.SpacingBefore = paragraph.LineSpacingBeforePoints.Value;
            } else if (GetNativeSpacingBeforePoints(directSpacing, fontSize, lineHeight) is { } directSpacingBefore) {
                style.SpacingBefore = directSpacingBefore;
            } else if (styleDefaults.SpacingBefore.HasValue) {
                style.SpacingBefore = styleDefaults.SpacingBefore.Value;
            } else if (nativeDefaults.ParagraphSpacingBeforeDeclared) {
                style.SpacingBefore = nativeDefaults.ParagraphSpacingBefore;
            }

            if (paragraph.LineSpacingAfterPoints.HasValue) {
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

            if (itemSpacingDeclared) {
                style.ItemSpacing = style.SpacingAfter;
            }

            style.KeepTogether = ReadNativeDirectParagraphOnOff<W.KeepLines>(paragraph) ?? styleDefaults.KeepTogether ?? false;
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
        }

        private static string ResolveNativeInlineListMarkerSuffix(WordListLevelSuffix? suffix) =>
            WordDocumentTraversal.ResolveTextListMarkerSuffix(suffix);

        private static PdfCore.PdfTextRun CreateNativeListMarkerTextRun(
            string marker,
            WordParagraph paragraph,
            NativeResolvedTextStyle textStyle,
            NativeFontMap? nativeFontMap,
            bool includeSuffix = true) {
            WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraph);
            if (info == null) {
                return new PdfCore.PdfTextRun(marker + (includeSuffix ? " " : string.Empty), bold: textStyle.Bold, color: textStyle.Color,
                    italic: textStyle.Italic, fontSize: textStyle.FontSize, font: textStyle.Font,
                    fontFamily: textStyle.FontFamily);
            }

            return new PdfCore.PdfTextRun(
                marker + (includeSuffix ? ResolveNativeInlineListMarkerSuffix(info.Value.LevelSuffix) : string.Empty),
                bold: info.Value.MarkerBold ?? textStyle.Bold,
                color: ParseNativeColor(info.Value.MarkerColorHex) ?? textStyle.Color,
                italic: info.Value.MarkerItalic ?? textStyle.Italic,
                fontSize: info.Value.MarkerFontSize ?? textStyle.FontSize,
                font: ResolveNativeListMarkerFont(info.Value, marker, textStyle),
                fontFamily: nativeFontMap != null
                    ? ResolveNativeListMarkerFontFamily(info.Value, marker, textStyle, nativeFontMap)
                    : textStyle.FontFamily);
        }

        private static (double MarkerWidth, double MarkerGap) ResolveNativeListMarkerSpacing(W.LevelSuffixValues? levelSuffix, double markerTextWidth, double fontSize, double textIndent, double markerIndent) {
            if (levelSuffix == W.LevelSuffixValues.Nothing) {
                return (markerTextWidth, 0D);
            }

            if (levelSuffix == W.LevelSuffixValues.Space) {
                return (markerTextWidth, EstimateNativeListMarkerWidth(" ", fontSize));
            }

            double markerColumnWidth = Math.Max(0D, textIndent - markerIndent);
            double markerWidth = Math.Max(markerTextWidth, markerColumnWidth);
            double markerGap = Math.Max(0D, markerColumnWidth - markerWidth);
            return (markerWidth, markerGap);
        }

        private static (double MarkerWidth, double MarkerGap) ResolveNativeListMarkerSpacing(WordListLevelSuffix? levelSuffix, double markerTextWidth, double fontSize, double textIndent, double markerIndent) =>
            ResolveNativeListMarkerSpacing(levelSuffix.ToOpenXml(), markerTextWidth, fontSize, textIndent, markerIndent);

        private static double EstimateNativeListMarkerWidth(string marker, double fontSize) {
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

            return width;
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
                   DoubleEquals(left.LeftIndent, right.LeftIndent) &&
                   NullableDoubleEquals(left.MarkerGap, right.MarkerGap) &&
                   NullableDoubleEquals(left.MarkerWidth, right.MarkerWidth) &&
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
                   left.KeepWithNext == right.KeepWithNext;
        }

        private static bool NullableDoubleEquals(double? left, double? right) {
            if (left.HasValue != right.HasValue) {
                return false;
            }

            return !left.HasValue || DoubleEquals(left.Value, right!.Value);
        }

        private static bool DoubleEquals(double left, double right) =>
            Math.Abs(left - right) < 0.001D;

        private static PdfCore.PdfStandardFont? ResolveNativeListMarkerFont(WordDocumentTraversal.ListInfo info, string marker, NativeResolvedTextStyle markerTextStyle) {
            if (ShouldUseNativeListTextFontForNormalizedMarker(info, marker)) {
                return markerTextStyle.Font;
            }

            return PdfCore.PdfStandardFontMapper.TryMapFontFamily(info.MarkerFontFamily, out PdfCore.PdfStandardFont markerFont)
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

        private static List<WordParagraph> GetNativeRuns(WordParagraph paragraph) {
            if (paragraph._paragraph == null) {
                return new List<WordParagraph>();
            }

            var runs = new List<WordParagraph>();
            var fieldVisibility = WordComplexFieldRunVisibility.ForParagraph(paragraph._paragraph);
            AddNativeVisibleRuns(runs, paragraph, paragraph._paragraph, null, fieldVisibility);

            return runs;
        }

        private static void AddNativeVisibleRuns(List<WordParagraph> runs, WordParagraph paragraph,
            DocumentFormat.OpenXml.OpenXmlCompositeElement container, W.Hyperlink? hyperlink,
            WordComplexFieldRunVisibility fieldVisibility) {
            foreach (var element in container.ChildElements) {
                if (element is W.Run sourceRun) {
                    W.Run? visibleRun = fieldVisibility.GetVisibleRun(sourceRun, out var visibleSourceChildren);
                    if (visibleRun != null)
                        runs.Add(new WordParagraph(paragraph._document, paragraph._paragraph!, sourceRun) {
                            _hyperlink = hyperlink,
                            _visibleRun = ReferenceEquals(visibleRun, sourceRun) ? null : visibleRun,
                            _visibleRunSourceChildren = visibleSourceChildren
                        });
                } else if (element is W.Hyperlink nestedHyperlink) {
                    AddNativeVisibleRuns(runs, paragraph, nestedHyperlink, nestedHyperlink, fieldVisibility);
                } else if (element is W.SimpleField simpleField) {
                    if (fieldVisibility.IsVisible)
                        AddNativeVisibleRuns(runs, paragraph, simpleField, hyperlink, fieldVisibility);
                    else
                        fieldVisibility.ObserveDescendantRuns(simpleField);
                } else if (element is W.SdtRun sdtRun) {
                    if (IsNativeSimpleTextContentControl(sdtRun))
                        AddNativeSdtRunRuns(runs, paragraph, sdtRun, hyperlink, fieldVisibility);
                    else
                        fieldVisibility.ObserveDescendantRuns(sdtRun);
                } else if (element is W.CustomXmlRun customXml) {
                    AddNativeVisibleRuns(runs, paragraph, customXml, hyperlink, fieldVisibility);
                }
            }
        }

        private static void AddNativeSdtRunRuns(List<WordParagraph> runs, WordParagraph paragraph,
            W.SdtRun sdtRun, W.Hyperlink? hyperlink, WordComplexFieldRunVisibility fieldVisibility) {
            if (sdtRun.Descendants<W.FieldChar>().Any()) {
                if (sdtRun.SdtContentRun != null)
                    AddNativeVisibleRuns(runs, paragraph, sdtRun.SdtContentRun, hyperlink, fieldVisibility);
                return;
            }

            if (!fieldVisibility.IsVisible) return;
            if (TryGetNativeSdtRunPropertyValue(paragraph._document, sdtRun, out string? propertyValue)) {
                W.Run resolvedRun = CreateNativeResolvedSdtRun(sdtRun, propertyValue!);
                runs.Add(new WordParagraph(paragraph._document, paragraph._paragraph!, resolvedRun) { _hyperlink = hyperlink });
                return;
            }

            if (sdtRun.SdtContentRun != null)
                AddNativeVisibleRuns(runs, paragraph, sdtRun.SdtContentRun, hyperlink, fieldVisibility);
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
