using System;
using System.Collections.Generic;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static bool AddTextRunBlock(
            WordParagraph paragraph,
            OfficeFontInfo font,
            double lineHeight,
            double height,
            WordParagraphSpacing spacing,
            WordParagraphSpacingState spacingState,
            WordImageTextLayout textLayout,
            WordImageListMarker? listMarker,
            A.ColorScheme? colorScheme,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics) {
            if (!EnsureVerticalSpace(context, spacing.Before + height, diagnostics)) {
                return false;
            }

            context.Y += spacing.Before;
            if (context.IsTargetPage) {
                AddParagraphFrame(paragraph, context, textLayout, height, colorScheme);
                if (listMarker.HasValue) {
                    WordImageListMarker marker = listMarker.Value;
                    context.Drawing.AddText(
                        marker.Marker,
                        textLayout.MarkerLeft,
                        context.Y,
                        textLayout.MarkerWidth,
                        lineHeight,
                        marker.Font,
                        marker.Color,
                        marker.Alignment,
                        lineHeight,
                        wrapText: false);
                }

                string text = ResolveImageExportText(paragraph, context);
                context.Drawing.AddText(
                    text,
                    textLayout.TextLeft,
                    context.Y,
                    textLayout.TextWidth,
                    height,
                    font,
                    ResolveParagraphTextColor(paragraph, colorScheme),
                    MapTextAlignment(paragraph.ParagraphAlignment),
                    lineHeight,
                    wrapText: true,
                    padding: textLayout.Padding,
                    paragraphIndent: textLayout.ParagraphIndent);
            }

            context.Y += height + spacing.After;
            context.SetParagraphSpacingState(spacingState);
            return true;
        }

        private static bool AddPaginatedTextRun(
            WordParagraph paragraph,
            IReadOnlyList<string> lines,
            OfficeFontInfo font,
            double lineHeight,
            WordParagraphSpacing spacing,
            WordParagraphSpacingState spacingState,
            WordImageListMarker? listMarker,
            A.ColorScheme? colorScheme,
            bool avoidWidowAndOrphan,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics) {
            if (lines.Count == 0) {
                return false;
            }

            int lineIndex = 0;
            bool spacingBeforePending = true;
            bool renderedOnTargetPage = false;
            bool consumedText = false;

            while (lineIndex < lines.Count) {
                if (spacingBeforePending) {
                    if (ShouldAdvanceBeforeSpacing(context, spacing.Before, lineHeight)) {
                        context.AdvanceColumnOrPage();
                        if (context.PastTargetPage) {
                            return renderedOnTargetPage;
                        }

                        continue;
                    }

                    context.Y += spacing.Before;
                    spacingBeforePending = false;
                }

                double available = Math.Max(0D, context.ContentBottom - context.Y);
                int linesOnPage = (int)Math.Floor(available / lineHeight);
                if (linesOnPage <= 0) {
                    if (context.Y > context.Top && context.CanAdvancePageForOverflow) {
                        context.AdvanceColumnOrPage();
                        if (context.PastTargetPage) {
                            return renderedOnTargetPage;
                        }

                        continue;
                    }

                    if (!context.StoppedForPagination) {
                        AddDiagnostic(diagnostics, context.OverflowDiagnosticCode, context.OverflowDiagnosticMessage);
                        context.StoppedForPagination = true;
                    }

                    return renderedOnTargetPage;
                }

                int lineCount = Math.Min(linesOnPage, lines.Count - lineIndex);
                lineCount = ApplyWidowOrphanControl(lines.Count, lineIndex, lineCount, avoidWidowAndOrphan, context);
                if (lineCount <= 0) {
                    if (context.CanAdvancePageForOverflow) {
                        context.AdvanceColumnOrPage();
                        if (context.PastTargetPage) {
                            return renderedOnTargetPage;
                        }

                        continue;
                    }

                    if (!context.StoppedForPagination) {
                        AddDiagnostic(diagnostics, context.OverflowDiagnosticCode, context.OverflowDiagnosticMessage);
                        context.StoppedForPagination = true;
                    }

                    return renderedOnTargetPage;
                }

                WordImageListMarker? currentMarker = lineIndex == 0 ? listMarker : null;
                WordImageTextLayout textLayout = ResolveTextLayout(context, currentMarker, paragraph);
                double sliceHeight = Math.Max(lineHeight, lineCount * lineHeight);
                if (context.IsTargetPage) {
                    AddTextRunSlice(paragraph, lines, lineIndex, lineCount, font, lineHeight, textLayout, currentMarker, colorScheme, context);
                    renderedOnTargetPage = true;
                }

                context.Y += sliceHeight;
                consumedText = true;
                lineIndex += lineCount;

                if (lineIndex < lines.Count) {
                    context.AdvanceColumnOrPage();
                    if (context.PastTargetPage) {
                        return renderedOnTargetPage;
                    }
                }
            }

            context.Y += spacing.After;
            context.SetParagraphSpacingState(spacingState);
            return consumedText || renderedOnTargetPage;
        }

        private static bool AddPaginatedRichTextRun(
            WordParagraph paragraph,
            IReadOnlyList<OfficeRichTextRun> richRuns,
            double maxFontSize,
            double lineHeight,
            WordParagraphSpacing spacing,
            WordParagraphSpacingState spacingState,
            WordImageListMarker? listMarker,
            A.ColorScheme? colorScheme,
            bool avoidWidowAndOrphan,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics) {
            WordImageTextLayout initialTextLayout = ResolveTextLayout(context, listMarker, paragraph);
            OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                richRuns,
                initialTextLayout.ContentWidth,
                double.MaxValue,
                Math.Max(1D, lineHeight / Math.Max(1D, maxFontSize)),
                CreateRichTextMeasure(context.CancellationToken),
                wrap: true,
                shrinkToFit: false,
                minimumFontSize: Math.Min(6D, maxFontSize),
                overflowBehavior: OfficeTextOverflowBehavior.Clip,
                paragraphIndent: initialTextLayout.ParagraphIndent,
                cancellationToken: context.CancellationToken);
            IReadOnlyList<OfficeRichTextLine> lines = layout.Lines;
            if (lines.Count == 0) {
                return false;
            }

            int lineIndex = 0;
            bool spacingBeforePending = true;
            bool renderedOnTargetPage = false;
            bool consumedText = false;

            while (lineIndex < lines.Count) {
                if (spacingBeforePending) {
                    double firstLineHeight = ResolveRichTextSliceLineHeight(lines[lineIndex], lineHeight);
                    if (ShouldAdvanceBeforeSpacing(context, spacing.Before, firstLineHeight)) {
                        context.AdvanceColumnOrPage();
                        if (context.PastTargetPage) {
                            return renderedOnTargetPage;
                        }

                        continue;
                    }

                    context.Y += spacing.Before;
                    spacingBeforePending = false;
                }

                double available = Math.Max(0D, context.ContentBottom - context.Y);
                int lineCount = CountRichTextLinesForPage(lines, lineIndex, available, lineHeight, out double sliceHeight);
                lineCount = ApplyWidowOrphanControl(lines.Count, lineIndex, lineCount, avoidWidowAndOrphan, context);
                if (lineCount > 0) {
                    sliceHeight = CalculateRichTextSliceHeight(lines, lineIndex, lineCount, lineHeight);
                }

                if (lineCount <= 0) {
                    if (context.Y > context.Top && context.CanAdvancePageForOverflow) {
                        context.AdvanceColumnOrPage();
                        if (context.PastTargetPage) {
                            return renderedOnTargetPage;
                        }

                        continue;
                    }

                    if (!context.StoppedForPagination) {
                        AddDiagnostic(diagnostics, context.OverflowDiagnosticCode, context.OverflowDiagnosticMessage);
                        context.StoppedForPagination = true;
                    }

                    return renderedOnTargetPage;
                }

                WordImageListMarker? currentMarker = lineIndex == 0 ? listMarker : null;
                WordImageTextLayout textLayout = ResolveTextLayout(context, currentMarker, paragraph);
                if (context.IsTargetPage) {
                    AddRichTextRunSlice(paragraph, lines, lineIndex, lineCount, lineHeight, sliceHeight, textLayout, currentMarker, colorScheme, context);
                    renderedOnTargetPage = true;
                }

                context.Y += sliceHeight;
                consumedText = true;
                lineIndex += lineCount;

                if (lineIndex < lines.Count) {
                    context.AdvanceColumnOrPage();
                    if (context.PastTargetPage) {
                        return renderedOnTargetPage;
                    }
                }
            }

            context.Y += spacing.After;
            context.SetParagraphSpacingState(spacingState);
            return consumedText || renderedOnTargetPage;
        }

        private static bool ShouldAdvanceBeforeSpacing(WordImageFlowContext context, double spacingBefore, double firstLineHeight) =>
            context.Y > context.Top &&
            context.Y + spacingBefore + firstLineHeight > context.ContentBottom &&
            context.CanAdvancePageForOverflow;

        private static int ApplyWidowOrphanControl(
            int totalLineCount,
            int lineIndex,
            int lineCount,
            bool avoidWidowAndOrphan,
            WordImageFlowContext context) {
            if (!avoidWidowAndOrphan || lineCount <= 0 || totalLineCount < 3) {
                return lineCount;
            }

            int remainingLineCount = totalLineCount - lineIndex;
            if (remainingLineCount <= lineCount) {
                return lineCount;
            }

            if (lineIndex == 0 && lineCount == 1 && context.Y > context.Top) {
                return 0;
            }

            int nextPageLineCount = remainingLineCount - lineCount;
            if (nextPageLineCount == 1 && lineCount > 1) {
                if (lineIndex == 0 && lineCount == 2 && context.Y > context.Top) {
                    return 0;
                }

                return lineCount - 1;
            }

            return lineCount;
        }

        private static void AddTextRunSlice(
            WordParagraph paragraph,
            IReadOnlyList<string> lines,
            int lineIndex,
            int lineCount,
            OfficeFontInfo font,
            double lineHeight,
            WordImageTextLayout textLayout,
            WordImageListMarker? listMarker,
            A.ColorScheme? colorScheme,
            WordImageFlowContext context) {
            double height = Math.Max(lineHeight, lineCount * lineHeight);
            AddParagraphFrame(paragraph, context, textLayout, height, colorScheme);
            if (listMarker.HasValue) {
                WordImageListMarker marker = listMarker.Value;
                context.Drawing.AddText(
                    marker.Marker,
                    textLayout.MarkerLeft,
                    context.Y,
                    textLayout.MarkerWidth,
                    lineHeight,
                    marker.Font,
                    marker.Color,
                    marker.Alignment,
                    lineHeight,
                    wrapText: false);
            }

            string text = string.Join(Environment.NewLine, CopyLineRange(lines, lineIndex, lineCount));
            context.Drawing.AddText(
                text,
                textLayout.TextLeft,
                context.Y,
                textLayout.TextWidth,
                height,
                font,
                ResolveParagraphTextColor(paragraph, colorScheme),
                MapTextAlignment(paragraph.ParagraphAlignment),
                lineHeight,
                wrapText: true,
                padding: textLayout.Padding,
                paragraphIndent: textLayout.ParagraphIndent);
        }

        private static List<string> CopyLineRange(IReadOnlyList<string> lines, int lineIndex, int lineCount) {
            var selected = new List<string>(lineCount);
            int lastIndex = Math.Min(lines.Count, lineIndex + lineCount);
            for (int i = lineIndex; i < lastIndex; i++) {
                selected.Add(lines[i]);
            }

            return selected;
        }

        private static void AddRichTextRunSlice(
            WordParagraph paragraph,
            IReadOnlyList<OfficeRichTextLine> lines,
            int lineIndex,
            int lineCount,
            double lineHeight,
            double height,
            WordImageTextLayout textLayout,
            WordImageListMarker? listMarker,
            A.ColorScheme? colorScheme,
            WordImageFlowContext context) {
            AddParagraphFrame(paragraph, context, textLayout, height, colorScheme);
            if (listMarker.HasValue) {
                WordImageListMarker marker = listMarker.Value;
                context.Drawing.AddText(
                    marker.Marker,
                    textLayout.MarkerLeft,
                    context.Y,
                    textLayout.MarkerWidth,
                    lineHeight,
                    marker.Font,
                    marker.Color,
                    marker.Alignment,
                    lineHeight,
                    wrapText: false);
            }

            context.Drawing.AddRichText(
                CreateRichTextRunsFromLines(lines, lineIndex, lineCount),
                textLayout.TextLeft,
                context.Y,
                textLayout.TextWidth,
                height,
                MapTextAlignment(paragraph.ParagraphAlignment),
                lineHeight,
                wrapText: true,
                padding: textLayout.Padding,
                paragraphIndent: textLayout.ParagraphIndent);
        }

        private static int CountRichTextLinesForPage(
            IReadOnlyList<OfficeRichTextLine> lines,
            int lineIndex,
            double available,
            double fallbackLineHeight,
            out double height) {
            height = 0D;
            int count = 0;
            for (int i = lineIndex; i < lines.Count; i++) {
                double lineHeight = ResolveRichTextSliceLineHeight(lines[i], fallbackLineHeight);
                if (count > 0 && height + lineHeight > available + 0.01D) {
                    break;
                }

                if (count == 0 && lineHeight > available + 0.01D) {
                    break;
                }

                height += lineHeight;
                count++;
            }

            return count;
        }

        private static double CalculateRichTextSliceHeight(
            IReadOnlyList<OfficeRichTextLine> lines,
            int lineIndex,
            int lineCount,
            double fallbackLineHeight) {
            double height = 0D;
            int lastIndex = Math.Min(lines.Count, lineIndex + lineCount);
            for (int i = lineIndex; i < lastIndex; i++) {
                height += ResolveRichTextSliceLineHeight(lines[i], fallbackLineHeight);
            }

            return height;
        }

        private static double ResolveRichTextSliceLineHeight(OfficeRichTextLine line, double fallbackLineHeight) =>
            line.LineHeight > 0D ? line.LineHeight : fallbackLineHeight;

        private static List<OfficeRichTextRun> CreateRichTextRunsFromLines(
            IReadOnlyList<OfficeRichTextLine> lines,
            int lineIndex,
            int lineCount) {
            var runs = new List<OfficeRichTextRun>();
            int lastIndex = Math.Min(lines.Count, lineIndex + lineCount);
            for (int i = lineIndex; i < lastIndex; i++) {
                OfficeRichTextLine line = lines[i];
                for (int segmentIndex = 0; segmentIndex < line.Segments.Count; segmentIndex++) {
                    OfficeRichTextSegment segment = line.Segments[segmentIndex];
                    runs.Add(new OfficeRichTextRun(
                        segment.Text,
                        segment.FontSize,
                        segment.Color,
                        segment.Bold,
                        segment.Italic,
                        segment.Underline,
                        segment.FontFamily,
                        segment.Strikethrough,
                        segment.BackgroundColor));
                }

                if (i < lastIndex - 1) {
                    OfficeRichTextSegment? style = line.Segments.Count > 0
                        ? line.Segments[line.Segments.Count - 1]
                        : FindNearestRichTextSegment(lines, i + 1, lastIndex);
                    if (style != null) {
                        runs.Add(new OfficeRichTextRun(
                            Environment.NewLine,
                            style.FontSize,
                            style.Color,
                            style.Bold,
                            style.Italic,
                            style.Underline,
                            style.FontFamily,
                            style.Strikethrough,
                            style.BackgroundColor));
                    }
                }
            }

            return runs;
        }

        private static OfficeRichTextSegment? FindNearestRichTextSegment(IReadOnlyList<OfficeRichTextLine> lines, int start, int end) {
            for (int i = start; i < end; i++) {
                if (lines[i].Segments.Count > 0) {
                    return lines[i].Segments[0];
                }
            }

            return null;
        }

        private static List<string> WrapTextIntoMeasuredLines(
            string text,
            OfficeFontInfo font,
            double contentWidth,
            CancellationToken cancellationToken = default,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint = null) {
            cancellationToken.ThrowIfCancellationRequested();
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            cancellationToken.ThrowIfCancellationRequested();
            string[] explicitLines = normalized.Split('\n');
            var lines = new List<string>();
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(font);
            OfficeTextMeasurementStyle style = measurer.CreateStyle(font, 72D);
            foreach (string explicitLine in explicitLines) {
                cancellationToken.ThrowIfCancellationRequested();
                AddMeasuredWrappedLine(
                    explicitLine,
                    Math.Max(1D, contentWidth),
                    measurer,
                    style,
                    lines,
                    cancellationToken,
                    cancellationCheckpoint);
            }

            return lines;
        }

        private static void AddMeasuredWrappedLine(
            string text,
            double contentWidth,
            OfficeTextMeasurer measurer,
            OfficeTextMeasurementStyle style,
            List<string> lines,
            CancellationToken cancellationToken,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint) {
            cancellationToken.ThrowIfCancellationRequested();
            if (text.Length == 0) {
                lines.Add(string.Empty);
                return;
            }

            List<string> words = SplitMeasuredWords(
                text,
                cancellationToken,
                cancellationCheckpoint);
            string current = string.Empty;
            for (int i = 0; i < words.Count; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                string word = words[i];
                if (current.Length == 0 &&
                    IsMeasuredWhitespaceRun(word, cancellationToken)) {
                    continue;
                }

                string candidate = current + word;
                if (measurer.MeasureWidth(candidate, style) <= contentWidth || current.Length == 0) {
                    if (measurer.MeasureWidth(candidate, style) <= contentWidth) {
                        current = candidate;
                        continue;
                    }

                    AddMeasuredWordFragments(
                        word,
                        contentWidth,
                        measurer,
                        style,
                        lines,
                        cancellationToken);
                    continue;
                }

                lines.Add(TrimMeasuredLine(current));
                current = string.Empty;
                i--;
            }

            if (current.Length > 0) {
                lines.Add(TrimMeasuredLine(current));
            }
        }

        private static string TrimMeasuredLine(string text) =>
            text.TrimEnd(' ', '\t');

        private static bool IsMeasuredWhitespaceRun(
            string text,
            CancellationToken cancellationToken) {
            for (int index = 0; index < text.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!char.IsWhiteSpace(text[index])) {
                    return false;
                }
            }
            return true;
        }

        private static List<string> SplitMeasuredWords(
            string text,
            CancellationToken cancellationToken,
            Action<WordImageCancellationCheckpoint>? cancellationCheckpoint) {
            var words = new List<string>();
            int start = -1;
            bool? currentIsWhitespace = null;
            for (int i = 0; i < text.Length; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (i == 1024) {
                    cancellationCheckpoint?.Invoke(
                        WordImageCancellationCheckpoint.PlainTextWrapping);
                    cancellationToken.ThrowIfCancellationRequested();
                }
                bool isWhitespace = char.IsWhiteSpace(text[i]);
                if (start >= 0 && currentIsWhitespace != isWhitespace) {
                    words.Add(text.Substring(start, i - start));
                    start = i;
                    currentIsWhitespace = isWhitespace;
                    continue;
                }

                if (start < 0) {
                    start = i;
                    currentIsWhitespace = isWhitespace;
                }
            }

            if (start >= 0) {
                words.Add(text.Substring(start));
            }

            if (words.Count == 0) {
                words.Add(string.Empty);
            }

            return words;
        }

        private static void AddMeasuredWordFragments(
            string word,
            double contentWidth,
            OfficeTextMeasurer measurer,
            OfficeTextMeasurementStyle style,
            List<string> lines,
            CancellationToken cancellationToken) {
            string current = string.Empty;
            for (int i = 0; i < word.Length; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                string candidate = current + word[i];
                if (current.Length > 0 && measurer.MeasureWidth(candidate, style) > contentWidth) {
                    lines.Add(current);
                    current = word[i].ToString();
                    continue;
                }

                current = candidate;
            }

            if (current.Length > 0) {
                lines.Add(current);
            }
        }
    }
}
