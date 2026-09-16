using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private const double DefaultListLeftIndentPoints = 36D;
        private const double DefaultListHangingIndentPoints = 18D;

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
                WordDocumentTraversal.ResolveTextListMarkerSuffix(info?.LevelSuffix),
                info?.PictureBulletId);
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

        private static OfficeRichTextRun CreateListMarkerRichTextRun(WordImageListMarker marker, double? availableWidth = null) =>
            CreateListMarkerRichTextRun(marker, marker.Marker + ResolveRichTextListMarkerSuffix(marker, availableWidth));

        private static OfficeRichTextRun CreateListMarkerRichTextRun(WordImageListMarker marker, string text) =>
            new OfficeRichTextRun(
                text, marker.Font.Size, marker.Color,
                marker.Font.IsBold, marker.Font.IsItalic, marker.Font.IsUnderline,
                marker.Font.FamilyName, marker.Font.IsStrikethrough);

        private static IReadOnlyList<OfficeRichTextRun> CreateAlignedListMarkerRichTextRuns(WordImageListMarker marker, double availableWidth) {
            if (marker.Alignment == OfficeTextAlignment.Left) {
                return new[] { CreateListMarkerRichTextRun(marker, availableWidth) };
            }

            (double _, double trailingWidth, double spaceWidth) = ResolveListMarkerAlignmentSpacing(marker, availableWidth);
            var runs = new List<OfficeRichTextRun>(2) {
                CreateListMarkerRichTextRun(marker, marker.Marker)
            };
            if (trailingWidth > 0.01D) runs.Add(CreateListMarkerSpacingRichTextRun(marker, trailingWidth, spaceWidth));
            return runs;
        }

        private static (double LeadingWidth, double TrailingWidth, double SpaceWidth) ResolveListMarkerAlignmentSpacing(WordImageListMarker marker, double availableWidth) {
            double textOffset = Math.Min(Math.Max(0D, marker.LeftIndentPoints), Math.Max(0D, availableWidth - 1D));
            double markerOffset = Math.Max(0D, textOffset - Math.Max(0D, marker.HangingIndentPoints));
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(marker.Font);
            OfficeTextMeasurementStyle markerStyle = measurer.CreateStyle(marker.Font, 72D);
            double markerWidth = measurer.MeasureWidth(marker.Marker, markerStyle);
            double markerColumnWidth = Math.Max(markerWidth, textOffset - markerOffset);
            double leadingWidth = marker.Alignment == OfficeTextAlignment.Right
                ? Math.Max(0D, markerColumnWidth - markerWidth)
                : Math.Max(0D, (markerColumnWidth - markerWidth) / 2D);
            double spaceWidth = Math.Max(0.01D, measurer.MeasureWidth(" ", markerStyle));
            double suffixWidth = marker.Suffix == " " ? spaceWidth : 0D;
            double trailingWidth = Math.Max(0D, markerColumnWidth - leadingWidth - markerWidth) + suffixWidth;
            return (leadingWidth, trailingWidth, spaceWidth);
        }

        private static OfficeRichTextRun CreateListMarkerSpacingRichTextRun(WordImageListMarker marker, double width, double spaceWidth) {
            int spaces = Math.Max(1, (int)Math.Ceiling(width / spaceWidth));
            spaces = Math.Min(8_192, spaces);
            double fontSize = Math.Max(1D, marker.Font.Size * width / (spaces * spaceWidth));
            var font = new OfficeFontInfo(marker.Font.FamilyName, fontSize, marker.Font.Style);
            return new OfficeRichTextRun(
                new string(' ', spaces), font.Size, marker.Color,
                font.IsBold, font.IsItalic, font.IsUnderline,
                font.FamilyName, font.IsStrikethrough);
        }

        private static string ResolveRichTextListMarkerSuffix(WordImageListMarker marker, double? availableWidth) {
            if (marker.Suffix != "\t" || !availableWidth.HasValue) {
                return marker.Suffix;
            }

            double textOffset = Math.Min(
                Math.Max(0D, marker.LeftIndentPoints),
                Math.Max(0D, availableWidth.Value - 1D));
            double markerOffset = Math.Max(0D, textOffset - Math.Max(0D, marker.HangingIndentPoints));
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(marker.Font);
            OfficeTextMeasurementStyle markerStyle = measurer.CreateStyle(marker.Font, 72D);
            double markerWidth = measurer.MeasureWidth(marker.Marker, markerStyle);
            double spaceWidth = Math.Max(0.01D, measurer.MeasureWidth(" ", markerStyle));
            double desiredGap = Math.Max(0D, textOffset - markerOffset - markerWidth);
            if (desiredGap <= 0D) {
                return string.Empty;
            }
            int spaces = Math.Max(1, (int)Math.Round(desiredGap / spaceWidth, MidpointRounding.AwayFromZero));
            return new string(' ', spaces);
        }

        private static WordImageTextLayout ResolveTextLayout(WordImageFlowContext context, WordImageListMarker? listMarker, WordParagraph? paragraph) {
            WordTextFlowFrame textFrame = context.ResolveTextFlowFrame();
            if (!listMarker.HasValue) {
                ResolveParagraphTextFrame(paragraph, textFrame.Width, out OfficeTextPadding padding, out OfficeTextParagraphIndent paragraphIndent);
                return new WordImageTextLayout(textFrame.Left, textFrame.Width, textFrame.Left, 0D, padding, paragraphIndent);
            }

            WordImageListMarker marker = listMarker.Value;
            double leftIndent = Math.Max(0D, marker.LeftIndentPoints);
            double hangingIndent = Math.Max(0D, marker.HangingIndentPoints);
            double textOffset = Math.Min(leftIndent, Math.Max(0D, textFrame.Width - 1D));
            double markerOffset = Math.Max(0D, textOffset - hangingIndent);
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(marker.Font);
            OfficeTextMeasurementStyle markerStyle = measurer.CreateStyle(marker.Font, 72D);
            double markerTextWidth = string.IsNullOrEmpty(marker.Marker)
                ? 0D
                : measurer.MeasureWidth(marker.Marker, markerStyle);
            double resolvedTextOffset;
            double markerWidth;
            if (marker.Marker.Length == 0) {
                markerWidth = 0D;
                resolvedTextOffset = textOffset;
            } else if (marker.Suffix.Length == 0) {
                bool useAlignedMarkerColumn = marker.Alignment != OfficeTextAlignment.Left;
                markerWidth = useAlignedMarkerColumn ? Math.Max(markerTextWidth, textOffset - markerOffset) : markerTextWidth;
                resolvedTextOffset = useAlignedMarkerColumn ? textOffset : markerOffset + markerTextWidth;
            } else if (marker.Suffix == " ") {
                bool useAlignedMarkerColumn = marker.Alignment != OfficeTextAlignment.Left;
                markerWidth = useAlignedMarkerColumn ? Math.Max(markerTextWidth, textOffset - markerOffset) : markerTextWidth;
                resolvedTextOffset = (useAlignedMarkerColumn ? textOffset : markerOffset + markerTextWidth) + measurer.MeasureWidth(" ", markerStyle);
            } else {
                markerWidth = Math.Max(markerTextWidth, textOffset - markerOffset);
                resolvedTextOffset = Math.Max(textOffset, markerOffset + markerTextWidth);
            }

            resolvedTextOffset = Math.Min(Math.Max(0D, resolvedTextOffset), Math.Max(0D, textFrame.Width - 1D));
            textOffset = Math.Min(Math.Max(0D, textOffset), Math.Max(0D, textFrame.Width - 1D));
            markerWidth = Math.Max(1D, Math.Min(markerWidth, Math.Max(1D, textFrame.Width - markerOffset)));
            double baseTextOffset = Math.Min(resolvedTextOffset, textOffset);
            double firstLineOffset = Math.Max(0D, resolvedTextOffset - baseTextOffset);
            double continuationLineOffset = Math.Max(0D, textOffset - baseTextOffset);
            OfficeTextParagraphIndent listParagraphIndent = firstLineOffset > 0D || continuationLineOffset > 0D
                ? new OfficeTextParagraphIndent(firstLineOffset, continuationLineOffset)
                : OfficeTextParagraphIndent.Empty;
            double textLeft = textFrame.Left + baseTextOffset;
            double textWidth = Math.Max(1D, textFrame.Width - baseTextOffset);

            return new WordImageTextLayout(textLeft, textWidth, textFrame.Left + markerOffset, markerWidth, OfficeTextPadding.Empty, listParagraphIndent);
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
                string suffix,
                int? pictureBulletId) {
                Marker = marker;
                Level = level;
                LeftIndentPoints = leftIndentPoints;
                HangingIndentPoints = hangingIndentPoints;
                Font = font;
                Color = color;
                Alignment = alignment;
                Suffix = suffix;
                PictureBulletId = pictureBulletId;
            }

            internal string Marker { get; }

            internal int Level { get; }

            internal double LeftIndentPoints { get; }

            internal double HangingIndentPoints { get; }

            internal OfficeFontInfo Font { get; }

            internal OfficeColor Color { get; }

            internal OfficeTextAlignment Alignment { get; }

            internal string Suffix { get; }

            internal int? PictureBulletId { get; }
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
