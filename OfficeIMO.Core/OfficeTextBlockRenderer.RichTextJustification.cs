using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextBlockRenderer {
    private static StringBuilder AppendSvgJustifiedRichTextLine(
        this StringBuilder builder,
        OfficeRichTextLine line,
        double x,
        double baseline,
        double textLength,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY) {
        OfficeRichTextSegment first = line.Segments[0];
        double backgroundCursor = x;
        for (int i = 0; i < line.Segments.Count; i++) {
            OfficeRichTextSegment segment = line.Segments[i];
            builder.AppendSvgRichTextSegmentBackground(segment, backgroundCursor, baseline, rotationDegrees, rotationCenterX, rotationCenterY);
            backgroundCursor += segment.Width;
        }

        builder.Append("<text")
            .AppendNumberAttribute("x", x)
            .AppendNumberAttribute("y", baseline)
            .AppendAttribute("font-family", string.IsNullOrWhiteSpace(first.FontFamily) ? "Arial, sans-serif" : first.FontFamily)
            .AppendNumberAttribute("font-size", first.FontSize)
            .AppendAttribute("text-anchor", "start")
            .AppendNumberAttribute("textLength", textLength)
            .AppendAttribute("lengthAdjust", "spacing");

        if (RequiresSvgWhitespacePreserve(line)) {
            builder.Append(" xml:space=\"preserve\"");
        }

        if (Math.Abs(rotationDegrees) > 0.000001D) {
            builder.AppendRotateTransformAttribute(rotationDegrees, rotationCenterX, rotationCenterY);
        }

        builder.Append('>');
        for (int i = 0; i < line.Segments.Count; i++) {
            OfficeRichTextSegment segment = line.Segments[i];
            builder.Append("<tspan")
                .AppendAttribute("font-family", string.IsNullOrWhiteSpace(segment.FontFamily) ? "Arial, sans-serif" : segment.FontFamily)
                .AppendNumberAttribute("font-size", segment.FontSize)
                .AppendPaintAttribute("fill", segment.Color);

            if (segment.Bold) {
                builder.Append(" font-weight=\"700\"");
            }

            if (segment.Italic) {
                builder.Append(" font-style=\"italic\"");
            }

            AppendSvgTextDecorationAttribute(builder, segment.Underline, segment.Strikethrough);
            builder.Append('>')
                .Append(OfficeSvgFormatting.Escape(segment.Text))
                .Append("</tspan>");
        }

        builder.Append("</text>");
        return builder;
    }

    private static bool ShouldJustifyRichTextLine(OfficeRichTextLine line, int lineIndex, int lineCount, double availableWidth, OfficeTextAlignment alignment) {
        return alignment == OfficeTextAlignment.Justify &&
            lineIndex < lineCount - 1 &&
            availableWidth > line.Width + 0.01D &&
            CountRichTextJustifiableWords(line) > 1;
    }

    private static int CountRichTextJustifiableWords(OfficeRichTextLine line) {
        int words = 0;
        bool insideWord = false;
        for (int segmentIndex = 0; segmentIndex < line.Segments.Count; segmentIndex++) {
            string text = line.Segments[segmentIndex].Text;
            for (int i = 0; i < text.Length; i++) {
                if (char.IsWhiteSpace(text[i])) {
                    insideWord = false;
                } else if (!insideWord) {
                    words++;
                    insideWord = true;
                }
            }
        }

        return words;
    }

    private static void DrawRasterJustifiedRichTextLine(
        OfficeRasterCanvas canvas,
        OfficeRichTextLine line,
        double left,
        double availableWidth,
        double baseline,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool flipHorizontal,
        bool flipVertical) {
        List<RichTextRenderToken> tokens = CreateRichTextRenderTokens(line, canvas);
        int gapCount = CountJustifiableRichTextGaps(tokens);
        if (gapCount == 0) {
            return;
        }

        double extraGap = Math.Max(0D, (availableWidth - line.Width) / gapCount);
        bool hasWordBefore = false;
        double cursor = left;
        for (int i = 0; i < tokens.Count; i++) {
            RichTextRenderToken token = tokens[i];
            if (token.IsWhitespace) {
                double whitespaceWidth = token.Width;
                if (hasWordBefore && HasWordAfter(tokens, i + 1)) {
                    whitespaceWidth += extraGap;
                }

                DrawRasterRichTextSegmentTokenBackground(
                    canvas,
                    token.Segment,
                    cursor,
                    baseline - (token.Segment.FontSize * 0.84D),
                    whitespaceWidth,
                    rotationDegrees,
                    rotationCenterX,
                    rotationCenterY,
                    flipHorizontal,
                    flipVertical);
                cursor += token.Width;
                if (hasWordBefore && HasWordAfter(tokens, i + 1)) {
                    cursor += extraGap;
                }

                continue;
            }

            double segmentTop = baseline - (token.Segment.FontSize * 0.84D);
            DrawRasterRichTextSegmentTokenBackground(
                canvas,
                token.Segment,
                cursor,
                segmentTop,
                token.Width,
                rotationDegrees,
                rotationCenterX,
                rotationCenterY,
                flipHorizontal,
                flipVertical);
            canvas.DrawTextLine(
                token.Text,
                cursor,
                segmentTop,
                token.Segment.FontSize,
                token.Segment.Color,
                token.Segment.Bold,
                token.Segment.Italic,
                OfficeTextAlignment.Left,
                rotationDegrees,
                rotationCenterX,
                rotationCenterY,
                token.Segment.Underline,
                token.Segment.Strikethrough,
                token.Segment.FontFamily,
                flipHorizontal,
                flipVertical);
            cursor += token.Width;
            hasWordBefore = true;
        }
    }

    private static List<RichTextRenderToken> CreateRichTextRenderTokens(OfficeRichTextLine line, OfficeRasterCanvas canvas) {
        var tokens = new List<RichTextRenderToken>();
        for (int segmentIndex = 0; segmentIndex < line.Segments.Count; segmentIndex++) {
            OfficeRichTextSegment segment = line.Segments[segmentIndex];
            string text = segment.Text;
            int tokenStart = 0;
            bool? whitespace = null;
            for (int i = 0; i < text.Length; i++) {
                bool currentWhitespace = char.IsWhiteSpace(text[i]);
                if (whitespace.HasValue && whitespace.Value != currentWhitespace) {
                    AddRichTextRenderToken(tokens, segment, text.Substring(tokenStart, i - tokenStart), whitespace.Value, canvas);
                    tokenStart = i;
                }

                whitespace = currentWhitespace;
            }

            if (whitespace.HasValue) {
                AddRichTextRenderToken(tokens, segment, text.Substring(tokenStart), whitespace.Value, canvas);
            }
        }

        return tokens;
    }

    private static void AddRichTextRenderToken(List<RichTextRenderToken> tokens, OfficeRichTextSegment segment, string text, bool whitespace, OfficeRasterCanvas canvas) {
        if (text.Length == 0) {
            return;
        }

        tokens.Add(new RichTextRenderToken(segment, text, canvas.MeasureText(text, segment.FontSize, segment.FontFamily), whitespace));
    }

    private static int CountJustifiableRichTextGaps(List<RichTextRenderToken> tokens) {
        int gaps = 0;
        bool hasWordBefore = false;
        for (int i = 0; i < tokens.Count; i++) {
            if (tokens[i].IsWhitespace) {
                if (hasWordBefore && HasWordAfter(tokens, i + 1)) {
                    gaps++;
                }
            } else {
                hasWordBefore = true;
            }
        }

        return gaps;
    }

    private static bool HasWordAfter(List<RichTextRenderToken> tokens, int startIndex) {
        for (int i = startIndex; i < tokens.Count; i++) {
            if (!tokens[i].IsWhitespace) {
                return true;
            }
        }

        return false;
    }

    private static bool RequiresSvgWhitespacePreserve(OfficeRichTextLine line) {
        for (int i = 0; i < line.Segments.Count; i++) {
            if (RequiresSvgWhitespacePreserve(line.Segments[i].Text)) {
                return true;
            }
        }

        return false;
    }

    private readonly struct RichTextRenderToken {
        internal RichTextRenderToken(OfficeRichTextSegment segment, string text, double width, bool isWhitespace) {
            Segment = segment;
            Text = text;
            Width = width;
            IsWhitespace = isWhitespace;
        }

        internal OfficeRichTextSegment Segment { get; }

        internal string Text { get; }

        internal double Width { get; }

        internal bool IsWhitespace { get; }
    }

    private static void DrawRasterRichTextSegmentTokenBackground(
        OfficeRasterCanvas canvas,
        OfficeRichTextSegment segment,
        double x,
        double top,
        double width,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool flipHorizontal,
        bool flipVertical) {
        if (!segment.BackgroundColor.HasValue || segment.BackgroundColor.Value.A == 0 || width <= 0D || segment.FontSize <= 0D) {
            return;
        }

        OfficeRichTextSegment backgroundSegment = new OfficeRichTextSegment(
            segment.Text,
            width,
            segment.FontSize,
            segment.Color,
            segment.Bold,
            segment.Italic,
            segment.Underline,
            segment.FontFamily,
            segment.Strikethrough,
            segment.BackgroundColor);
        DrawRasterRichTextSegmentBackground(canvas, backgroundSegment, x, top, rotationDegrees, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
    }
}
