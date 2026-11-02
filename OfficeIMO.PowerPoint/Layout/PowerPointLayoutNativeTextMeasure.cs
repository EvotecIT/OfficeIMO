using System.Globalization;
using OfficeIMO.Drawing;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Measures collected native text with the styles that binding preserves. Only missing properties use
    /// layout defaults; authored sizes, font families, tracking, margins and paragraph spacing remain effective.
    /// </summary>
    internal static class PowerPointLayoutNativeTextMeasure {
        internal static double MaximumSize(PowerPointParagraph paragraph, double fallback) =>
            paragraph.InlineNodes.Where(n => n.Run != null)
                .Select(n => PowerPointEffectiveRunStyleResolver.Resolve(n.Run!, paragraph, null, null).FontSizePoints ?? fallback)
                .DefaultIfEmpty(PowerPointEffectiveRunStyleResolver.ResolveParagraph(paragraph, null, null).FontSizePoints ?? fallback).Max();

        internal static double Height(PowerPointLayoutContext context, PowerPointParagraph paragraph, double size,
            double width, double margin = 0D, double? spaceAfter = null, int start = 0, int? length = null,
            PowerPointTextRole role = PowerPointTextRole.Body) {
            var style = context.Style(role);
            // Use the same font metrics as the shared Drawing renderer. Family factors are sufficient
            // for theme estimates, but can under-measure an authored face at a large native size.
            var metrics = OfficeDrawingTextLayout.CreateMetrics((OfficeFontFaceCollection?)null);
            double available = Math.Max(1D, (width - (paragraph.LeftMargin?.ToPoints() ?? margin)
                - (paragraph.Paragraph.ParagraphProperties?.RightMargin?.Value ?? 0) / 12700D
                - Math.Max(0D, paragraph.Indent?.ToPoints() ?? 0D)) / 1.06D);
            double height = 0D, lineWidth = 0D, lineSize = 0D, wordWidth = 0D, wordSize = 0D, spaces = 0D;
            int position = 0, end = start + (length ?? paragraph.Text.Length);
            bool lineStarted = false;
            double maximumSize = MaximumSize(paragraph, size);
            // Binding supplies point leading from the paragraph's largest effective run when no leading is authored.
            double Advance(double points) => Math.Max(points, paragraph.LineSpacing?.ToPoints()
                ?? (paragraph.LineSpacingMultiplier.HasValue ? points * paragraph.LineSpacingMultiplier.Value : maximumSize * style.LineHeight));
            void Line() {
                height += Advance(lineSize > 0D ? lineSize : size);
                lineWidth = lineSize = 0D;
                lineStarted = false;
            }
            void Word() {
                if (wordSize == 0D) return;
                if (lineStarted && lineWidth + spaces + wordWidth > available) { Line(); spaces = 0D; }
                lineWidth += (lineStarted ? spaces : 0D) + wordWidth;
                lineSize = Math.Max(lineSize, wordSize);
                lineStarted = true;
                // DrawingML wraps a token wider than its frame. Count all of its physical lines.
                while (lineWidth > available) { height += Advance(lineSize); lineWidth -= available; }
                wordWidth = wordSize = spaces = 0D;
            }
            foreach (var inline in paragraph.InlineNodes) {
                string text = inline.Text;
                int inlineStart = position;
                int from = Math.Max(start, inlineStart), to = Math.Min(end, inlineStart + text.Length);
                position += text.Length;
                if (to <= from) continue;
                text = text.Substring(from - inlineStart, to - from).Replace("\r\n", "\n").Replace('\r', '\n');
                var effective = inline.Run == null
                    ? PowerPointEffectiveRunStyleResolver.ResolveParagraph(paragraph, null, null)
                    : PowerPointEffectiveRunStyleResolver.Resolve(inline.Run, paragraph, null, null);
                double points = effective.FontSizePoints ?? size;
                string font = effective.FontName ?? context.FontOf(style);
                bool bold = effective.Bold ?? style.Bold;
                OfficeFontStyle fontStyle = bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular;
                if (effective.Italic == true) fontStyle |= OfficeFontStyle.Italic;
                double trackingPoints = inline.Run?.RunProperties?.Spacing?.Value / 100D
                    ?? paragraph.Paragraph.ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>()?.Spacing?.Value / 100D
                    ?? style.Tracking * size;
                // Drawing currently reports unsupported tracking. Reserve at least its untracked
                // advance, and also retain room for positive tracking in native Office.
                double Measure(string value) => metrics.MeasureText(value, points, font, fontStyle)
                    + Math.Max(0D, trackingPoints) * value.Length;
                foreach (Match token in Regex.Matches(text, @"\n|[^\S\n]+|[^\s]+")) {
                    if (token.Value == "\n") { Word(); Line(); spaces = 0D; }
                    else if (char.IsWhiteSpace(token.Value[0])) {
                        Word();
                        spaces += Measure(token.Value);
                    } else {
                        wordWidth += Measure(token.Value);
                        wordSize = Math.Max(wordSize, points);
                    }
                }
            }
            Word();
            Line();
            double markSize = MaximumSize(paragraph, size);
            double Spacing(DocumentFormat.OpenXml.OpenXmlCompositeElement? spacing, double fallback) =>
                spacing?.GetFirstChild<A.SpacingPoints>()?.Val?.Value / 100D
                ?? spacing?.GetFirstChild<A.SpacingPercent>()?.Val?.Value / 100000D * markSize
                ?? fallback;
            return height + Spacing(paragraph.Paragraph.ParagraphProperties?.SpaceBefore, 0D)
                + Spacing(paragraph.Paragraph.ParagraphProperties?.SpaceAfter, spaceAfter ?? context.Theme.ParagraphAfter);
        }

        /// <summary>Splits only when required, preferring sentences and words and preserving every source character.</summary>
        internal static IReadOnlyList<PowerPointParagraph> Split(PowerPointLayoutContext context, PowerPointParagraph paragraph,
            double size, double width, double height, double margin = 0D,
            PowerPointTextRole role = PowerPointTextRole.Body, double? spaceAfter = null) {
            if (Height(context, paragraph, size, width, margin, spaceAfter, role: role) <= height) return new[] { paragraph };
            string text = paragraph.Text;
            if (text.Length == 0) throw new ArgumentException("The authored paragraph spacing cannot fit a content frame.");
            int[] elements = StringInfo.ParseCombiningCharacters(text).Concat(new[] { text.Length }).ToArray();
            var pieces = new List<string>();
            int offset = 0;
            while (offset < text.Length) {
                int low = Array.BinarySearch(elements, offset) + 1, high = elements.Length - 1, best = offset;
                while (low <= high) {
                    int mid = low + (high - low) / 2;
                    if (Height(context, paragraph, size, width, margin, spaceAfter, start: offset, length: elements[mid] - offset, role: role) <= height) {
                        best = elements[mid]; low = mid + 1;
                    } else high = mid - 1;
                }
                if (best == offset) throw new ArgumentException("The authored native text size or spacing cannot fit a content frame.");
                if (best < text.Length) {
                    string candidate = text.Substring(offset, best - offset);
                    int sentence = Regex.Matches(candidate, @"(?<=[.!?])\s+").Cast<Match>().Select(m => m.Index + m.Length).DefaultIfEmpty(0).Last();
                    int word = Regex.Matches(candidate, @"\s+").Cast<Match>().Select(m => m.Index + m.Length).DefaultIfEmpty(0).Last();
                    best = offset + (sentence > 0 ? sentence : word > 0 ? word : candidate.Length);
                }
                pieces.Add(text.Substring(offset, best - offset));
                offset = best;
            }
            return PowerPointLayoutNativeText.SplitParagraph(paragraph, pieces);
        }
    }
    internal static partial class PowerPointLayoutEngine {
        /// <summary>Flows live text at its effective authored styles; continuation never rewrites native formatting.</summary>
        private static IEnumerable<PowerPointLayoutSlide> SplitNativeContent(PowerPointLayoutContext context, PowerPointLayoutSlide slide,
            PowerPointLayoutNativeText body, PowerPointLayoutList? list = null, PowerPointLayoutText? lede = null) {
            var style = context.Style(PowerPointTextRole.Body);
            double width = context.Grid.W(8);
            double firstHeight = ContentZone(context, MeasureHeader(context, slide)).Height
                - (lede != null ? NativeLedeHeight(context, lede, width, PowerPointTextRole.Lede, 64D) + 12D : 0D);
            double nextHeight = ContentZone(context, MeasureHeader(context, Continuation(context, slide, 2, 2))).Height;
            var chunks = NativeChunks(context, body, new PowerPointLayoutRect(0, 0, width, Math.Min(firstHeight, nextHeight)),
                list != null, list?.Numbered ?? false);
            int start = 1;
            for (int page = 0; page < chunks.Count; page++) {
                var part = page == 0 ? slide.CloneHeader() : Continuation(context, slide, page + 1, chunks.Count);
                if (page == 0) { part.Subtitle = slide.Subtitle; part.Notes = slide.Notes; if (lede != null) part.Blocks.Add(lede); }
                part.Layout = PowerPointLayout.TitleAndContent;
                part.ListSize = style.Size;
                part.ListStart = start;
                var slice = body.Slice(chunks[page]);
                part.Blocks.Add(list == null ? new PowerPointLayoutText(Array.Empty<string>()) { NativeBody = slice }
                    : new PowerPointLayoutList(Array.Empty<string>(), list.Numbered) { NativeBody = slice });
                start += chunks[page].Count(p => (p.Level ?? 0) == 0);
                yield return part;
            }
        }

        /// <summary>Fits each accepted block independently, preserving siblings on its first physical page.</summary>
        private static IEnumerable<PowerPointLayoutSlide> SplitNativeSpecial(PowerPointLayoutContext context, PowerPointLayoutSlide slide) {
            var pages = slide.Blocks.Select(block => SplitNativeSpecialBlock(context, slide, block)).ToList();
            int count = pages.Select(block => block.Count).DefaultIfEmpty(1).Max();
            int nextStep = slide.ListStart;
            for (int page = 0; page < count; page++) {
                // These compositions have no continuation header, so keep their header geometry stable.
                var part = slide.CloneHeader();
                part.Notes = page == 0 ? slide.Notes : null;
                part.Layout = slide.Layout;
                part.ListStart = nextStep;
                foreach (var blocks in pages) {
                    if (page >= blocks.Count) continue;
                    var block = blocks[page];
                    part.Blocks.Add(block);
                    if (block is PowerPointLayoutList list) nextStep += list.Items.Count(item => item.Level == 1);
                }
                yield return part;
            }
        }

        private static IReadOnlyList<PowerPointLayoutBlock> SplitNativeSpecialBlock(PowerPointLayoutContext context,
            PowerPointLayoutSlide slide, PowerPointLayoutBlock block) {
            var body = (block as PowerPointLayoutText)?.NativeBody ?? (block as PowerPointLayoutList)?.NativeBody;
            // Empty native lists still need their original block to reach the renderer and bind a valid live frame.
            if (body == null || body.Paragraphs.Count == 0) return new[] { block };
            bool closing = slide.Layout == PowerPointLayout.Closing, list = block is PowerPointLayoutList;
            var role = list ? PowerPointTextRole.Body : closing ? PowerPointTextRole.Subtitle : PowerPointTextRole.Quote;
            var box = closing ? context.Grid.Box(list ? 8 : 1, list ? 5 : 6, 3, 8) : context.Grid.Box(2, 9, 4, 6);
            double size = context.Style(role).Size;
            double overhead = 0D;
            if (closing && !list) {
                var stack = ClosingStack(context, slide, box.Width);
                double fixedHeight = stack.Where(item => item.Paragraph.NativeParagraph == null).Sum(item => item.Height)
                    + stack.Take(stack.Count - 1).Sum(item => item.GapAfter);
                box = box.WithHeight(box.Height - fixedHeight - 4D);
            } else if (closing) {
                var eyebrow = context.Style(PowerPointTextRole.Eyebrow);
                overhead = 12D + eyebrow.Size * eyebrow.LineHeight + 6D + 18D + 4D;
            }
            var chunks = new List<List<PowerPointParagraph>> { new() };
            double used = 0D;
            foreach (var paragraph in body.Paragraphs) {
                foreach (var piece in PowerPointLayoutNativeTextMeasure.Split(context, paragraph, size, box.Width,
                    box.Height - overhead, role: role, spaceAfter: 0D)) {
                    double height = PowerPointLayoutNativeTextMeasure.Height(context, piece, size, box.Width,
                        spaceAfter: 0D, role: role) + overhead;
                    if (chunks[chunks.Count - 1].Count > 0 && (!list || used + height > box.Height)) { chunks.Add(new()); used = 0D; }
                    chunks[chunks.Count - 1].Add(piece); used += height;
                }
            }
            return chunks.Select(chunk => list
                ? (PowerPointLayoutBlock)new PowerPointLayoutList(Array.Empty<string>(), ((PowerPointLayoutList)block).Numbered) { NativeBody = body.Slice(chunk) }
                : new PowerPointLayoutText(Array.Empty<string>()) { NativeBody = body.Slice(chunk) }).ToList();
        }

        private static IReadOnlyList<List<PowerPointParagraph>> NativeChunks(PowerPointLayoutContext context,
            PowerPointLayoutNativeText body, PowerPointLayoutRect box, bool list = false, bool numbered = false) {
            double size = context.Style(PowerPointTextRole.Body).Size;
            double Margin(PowerPointParagraph p) => !list ? 0D : numbered ? 28D
                : (p.Level ?? 0) == 0 ? context.Theme.ListIndent : context.Theme.ListIndent + 22D;
            double Size(PowerPointParagraph p) => !list || (p.Level ?? 0) == 0 ? size : Math.Round(size * 0.86D);
            var chunks = new List<List<PowerPointParagraph>> { new() };
            double used = 0D;
            foreach (var paragraph in body.Paragraphs) {
                foreach (var piece in PowerPointLayoutNativeTextMeasure.Split(context, paragraph, Size(paragraph), box.Width, box.Height, Margin(paragraph))) {
                    double height = PowerPointLayoutNativeTextMeasure.Height(context, piece, Size(piece), box.Width, Margin(piece));
                    if (chunks[chunks.Count - 1].Count > 0 && used + height > box.Height) { chunks.Add(new()); used = 0D; }
                    chunks[chunks.Count - 1].Add(piece); used += height;
                }
            }
            return chunks;
        }

        private static bool HasNativeText(PowerPointLayoutBlock block) => block is PowerPointLayoutText t && t.NativeBody != null
            || block is PowerPointLayoutList l && l.NativeBody != null
            || block is PowerPointLayoutColumn c && c.Blocks.Any(HasNativeText);

        /// <summary>Continues overflowing native columns using the same frame geometry as rendering.</summary>
        private static IEnumerable<PowerPointLayoutSlide> SplitNativeColumns(PowerPointLayoutContext context, PowerPointLayoutSlide slide,
            IReadOnlyList<PowerPointLayoutBlock> blocks, PowerPointLayoutText? lede) {
            var header = MeasureHeader(context, slide);
            var nextHeader = MeasureHeader(context, Continuation(context, slide, 2, 2));
            PowerPointLayoutRect[] Boxes(HeaderGeometry h) {
                if (slide.Layout == PowerPointLayout.Comparison) return new[] { ComparisonBodyBox(context, h, 0), ComparisonBodyBox(context, h, 1) };
                var zone = ContentZone(context, h);
                if (lede != null) zone = zone.Inset(0, NativeLedeHeight(context, lede, zone.Width, PowerPointTextRole.Lede, 64D) + 12D, 0, 0);
                return blocks.Count == 1 ? new[] { zone } : TwoContentBoxes(context, blocks, zone);
            }
            var firstBoxes = Boxes(header); var nextBoxes = Boxes(nextHeader);
            var pages = blocks.Select((b, index) => NativeBlockPages(context, b,
                firstBoxes[index].WithHeight(Math.Min(firstBoxes[index].Height, nextBoxes[index].Height)),
                slide.Layout != PowerPointLayout.Comparison)).ToList();
            int count = pages.Max(p => p.Count);
            for (int page = 0; page < count; page++) {
                var part = page == 0 ? slide.CloneHeader() : Continuation(context, slide, page + 1, count);
                if (page == 0) { part.Subtitle = slide.Subtitle; part.Notes = slide.Notes; if (lede != null) part.Blocks.Add(lede); }
                part.Layout = slide.Layout;
                part.Blocks.AddRange(pages.Select((p, index) => page < p.Count ? p[page]
                    : blocks[index] is PowerPointLayoutColumn column ? new PowerPointLayoutColumn(column.Title, Array.Empty<PowerPointLayoutBlock>())
                    : new PowerPointLayoutColumn(null, Array.Empty<PowerPointLayoutBlock>())));
                yield return part;
            }
        }

        private static IReadOnlyList<PowerPointLayoutBlock> NativeBlockPages(PowerPointLayoutContext context, PowerPointLayoutBlock block,
            PowerPointLayoutRect box, bool includeColumnTitle) {
            if (block is PowerPointLayoutColumn column) {
                var (lede, inner) = SplitLede(column.Blocks);
                if (includeColumnTitle && !string.IsNullOrWhiteSpace(column.Title))
                    box = box.Inset(0, context.Height(column.Title!, context.Style(PowerPointTextRole.Heading), box.Width) + 16D, 0, 0);
                if (lede != null) box = box.Inset(0, NativeLedeHeight(context, lede, box.Width, PowerPointTextRole.BodySmall,
                    context.Height(lede.Paragraphs[0], context.Style(PowerPointTextRole.BodySmall), box.Width)) + 10D, 0, 0);
                if (inner.Count == 0) return new[] { block };
                return NativeBlockPages(context, inner[0], box, false).Select((b, index) => (PowerPointLayoutBlock)new PowerPointLayoutColumn(column.Title,
                    (index == 0 && lede != null ? new PowerPointLayoutBlock[] { lede } : Array.Empty<PowerPointLayoutBlock>()).Concat(new[] { b }))).ToList();
            }
            var body = (block as PowerPointLayoutText)?.NativeBody ?? (block as PowerPointLayoutList)?.NativeBody;
            if (body == null) return new[] { block };
            var list = block as PowerPointLayoutList;
            var chunks = NativeChunks(context, body, box, list != null, list?.Numbered ?? false);
            var pages = chunks.Select(c => list == null
                ? (PowerPointLayoutBlock)new PowerPointLayoutText(Array.Empty<string>()) { NativeBody = body.Slice(c) }
                : new PowerPointLayoutList(Array.Empty<string>(), list.Numbered) { NativeBody = body.Slice(c) }).ToList();
            if (list?.Numbered == true) {
                int number = 1;
                for (int i = 0; i < pages.Count; i++) {
                    ((PowerPointLayoutList)pages[i]).NativeStart = number;
                    number += chunks[i].Count(p => (p.Level ?? 0) == 0);
                }
            }
            return pages;
        }
    }
}
