using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private const double DefaultListLeftIndentPoints = 36D;
        private const double DefaultListHangingIndentPoints = 18D;
        private const double ListMarkerGapPoints = 3D;

        private static WordImageListMarker? CreateListMarker(
            WordDocument document,
            Paragraph paragraph,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            WordParagraph paragraphReference = new WordParagraph(document, paragraph);
            if (!listMarkers.TryGetValue(paragraphReference, out (int Level, string Marker) marker)) {
                return null;
            }

            WordDocumentTraversal.ListInfo? info = WordDocumentTraversal.GetListInfo(paragraphReference);
            WordParagraph? firstRun = GetFirstTextRun(document, paragraph);
            OfficeFontInfo baseFont = firstRun == null ? OfficeFontInfo.Default : CreateFont(firstRun);
            OfficeFontInfo markerFont = CreateListMarkerFont(info, baseFont, WordDocumentTraversal.ShouldUseTextFontForMarker(info, marker.Marker));
            var colorScheme = GetDocumentColorScheme(document);
            OfficeColor markerColor = ResolveListMarkerColor(info, ResolveParagraphTextColor(firstRun, colorScheme));
            OfficeTextAlignment markerAlignment = info?.LevelJustification == WordListLevelAlignment.Right
                ? OfficeTextAlignment.Right
                : info?.LevelJustification == WordListLevelAlignment.Center
                    ? OfficeTextAlignment.Center
                    : OfficeTextAlignment.Left;

            return new WordImageListMarker(
                marker.Marker,
                marker.Level,
                ToPoints(info?.LeftIndentTwips, DefaultListLeftIndentPoints * (marker.Level + 1)),
                ToPoints(info?.HangingIndentTwips, DefaultListHangingIndentPoints),
                markerFont,
                markerColor,
                markerAlignment,
                WordDocumentTraversal.ResolveTextListMarkerSuffix(info?.LevelSuffix));
        }

        private static WordParagraph? GetFirstTextRun(WordDocument document, Paragraph paragraph) {
            foreach (WordParagraph run in WordSection.ConvertParagraphToWordParagraphs(document, paragraph, splitPaginationMarkers: true)) {
                if (!string.IsNullOrEmpty(run.Text)) {
                    return run;
                }
            }

            return null;
        }

        private static OfficeFontInfo CreateListMarkerFont(WordDocumentTraversal.ListInfo? info, OfficeFontInfo baseFont, bool useTextFont) {
            OfficeFontStyle style = OfficeFontStyle.Regular;
            if (info?.MarkerBold ?? baseFont.IsBold) {
                style |= OfficeFontStyle.Bold;
            }

            if (info?.MarkerItalic ?? baseFont.IsItalic) {
                style |= OfficeFontStyle.Italic;
            }

            return new OfficeFontInfo(
                useTextFont || string.IsNullOrWhiteSpace(info?.MarkerFontFamily) ? baseFont.FamilyName : info!.Value.MarkerFontFamily!,
                info?.MarkerFontSize ?? baseFont.Size,
                style);
        }

        private static OfficeColor ResolveListMarkerColor(WordDocumentTraversal.ListInfo? info, OfficeColor fallback) {
            string? colorHex = info?.MarkerColorHex;
            if (string.IsNullOrWhiteSpace(colorHex)) {
                return fallback;
            }

            try {
                return Helpers.ParseColor(colorHex!);
            } catch (ArgumentException) {
                return fallback;
            }
        }

        private static OfficeRichTextRun CreateListMarkerRichTextRun(WordImageListMarker marker) =>
            new OfficeRichTextRun(
                marker.Marker + marker.Suffix, marker.Font.Size, marker.Color,
                marker.Font.IsBold, marker.Font.IsItalic, marker.Font.IsUnderline,
                marker.Font.FamilyName, marker.Font.IsStrikethrough);

        private static WordImageTextLayout ResolveTextLayout(WordImageFlowContext context, WordImageListMarker? listMarker, WordParagraph? paragraph) {
            WordTextFlowFrame textFrame = context.ResolveTextFlowFrame();
            if (!listMarker.HasValue) {
                ResolveParagraphTextFrame(paragraph, textFrame.Width, out OfficeTextPadding padding, out OfficeTextParagraphIndent paragraphIndent);
                return new WordImageTextLayout(textFrame.Left, textFrame.Width, textFrame.Left, 0D, padding, paragraphIndent);
            }

            WordImageListMarker marker = listMarker.Value;
            double leftIndent = Math.Max(0D, marker.LeftIndentPoints);
            double hangingIndent = Math.Max(0D, marker.HangingIndentPoints);
            double textOffset = Math.Min(Math.Max(DefaultListHangingIndentPoints, leftIndent), Math.Max(DefaultListHangingIndentPoints, textFrame.Width - 1D));
            double markerOffset = Math.Max(0D, textOffset - hangingIndent);
            double markerWidth = Math.Max(1D, textOffset - markerOffset - ListMarkerGapPoints);
            double textLeft = textFrame.Left + textOffset;
            double textWidth = Math.Max(1D, textFrame.Width - textOffset);

            return new WordImageTextLayout(textLeft, textWidth, textFrame.Left + markerOffset, markerWidth, OfficeTextPadding.Empty, OfficeTextParagraphIndent.Empty);
        }

        private static void ResolveParagraphTextFrame(WordParagraph? paragraph, double contentWidth, out OfficeTextPadding padding, out OfficeTextParagraphIndent paragraphIndent) {
            if (paragraph == null) {
                padding = OfficeTextPadding.Empty;
                paragraphIndent = OfficeTextParagraphIndent.Empty;
                return;
            }

            double before = Math.Max(0D, paragraph.IndentationBeforePoints ?? 0D);
            double right = Math.Max(0D, paragraph.IndentationAfterPoints ?? 0D);
            double hanging = Math.Max(0D, paragraph.IndentationHangingPoints ?? 0D);
            double firstLine = Math.Max(0D, paragraph.IndentationFirstLinePoints ?? 0D);
            double left = hanging > 0D ? Math.Max(0D, before - hanging) : before;
            double maximumHorizontalPadding = Math.Max(0D, contentWidth - 1D);
            if (left + right > maximumHorizontalPadding) {
                left = Math.Min(left, maximumHorizontalPadding);
                right = Math.Min(right, Math.Max(0D, maximumHorizontalPadding - left));
            }

            double availableWidth = Math.Max(1D, contentWidth - left - right);
            if (hanging > 0D) {
                double continuationOffset = Math.Min(Math.Max(0D, before - left), Math.Max(0D, availableWidth - 1D));
                paragraphIndent = continuationOffset > 0D ? OfficeTextParagraphIndent.Hanging(continuationOffset) : OfficeTextParagraphIndent.Empty;
            } else {
                double firstLineOffset = Math.Min(firstLine, Math.Max(0D, availableWidth - 1D));
                paragraphIndent = firstLineOffset > 0D ? OfficeTextParagraphIndent.FirstLine(firstLineOffset) : OfficeTextParagraphIndent.Empty;
            }

            padding = new OfficeTextPadding(left, 0D, right, 0D);
        }

        private readonly struct WordImageListMarker {
            internal WordImageListMarker(
                string marker,
                int level,
                double leftIndentPoints,
                double hangingIndentPoints,
                OfficeFontInfo font,
                OfficeColor color,
                OfficeTextAlignment alignment,
                string suffix) {
                Marker = marker;
                Level = level;
                LeftIndentPoints = leftIndentPoints;
                HangingIndentPoints = hangingIndentPoints;
                Font = font;
                Color = color;
                Alignment = alignment;
                Suffix = suffix;
            }

            internal string Marker { get; }

            internal int Level { get; }

            internal double LeftIndentPoints { get; }

            internal double HangingIndentPoints { get; }

            internal OfficeFontInfo Font { get; }

            internal OfficeColor Color { get; }

            internal OfficeTextAlignment Alignment { get; }

            internal string Suffix { get; }
        }

        private readonly struct WordImageTextLayout {
            internal WordImageTextLayout(double textLeft, double textWidth, double markerLeft, double markerWidth, OfficeTextPadding padding, OfficeTextParagraphIndent paragraphIndent) {
                TextLeft = textLeft;
                TextWidth = textWidth;
                MarkerLeft = markerLeft;
                MarkerWidth = markerWidth;
                Padding = padding;
                ParagraphIndent = paragraphIndent;
            }

            internal double TextLeft { get; }

            internal double TextWidth { get; }

            internal double ContentWidth => Math.Max(1D, TextWidth - Padding.Horizontal);

            internal double LayoutWidth => Math.Max(1D, ContentWidth - ParagraphIndent.MaximumOffset);

            internal double MarkerLeft { get; }

            internal double MarkerWidth { get; }

            internal OfficeTextPadding Padding { get; }

            internal OfficeTextParagraphIndent ParagraphIndent { get; }
        }
    }
}
