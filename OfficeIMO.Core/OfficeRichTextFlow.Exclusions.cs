using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeRichTextFlow {
    /// <summary>
    /// Continues a story through all free intervals on each row. Fragments share
    /// a baseline; paragraph ends and source hard breaks finish the row. Masks
    /// cover the complete line extent, including larger later fragments.
    /// </summary>
    internal IReadOnlyList<OfficeRichTextFlowFragment> TakeAround(OfficeTextFlowRegion area,
        IReadOnlyList<OfficeTextFlowRegion> exclusions, OfficeDrawingTextMetrics metrics,
        Action accountInspection, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var result = new List<OfficeRichTextFlowFragment>();
        if (area.Width <= 0 || area.Height <= 0) return result;
        double top = area.Y;
        int spacedParagraph = -1;
        while (HasRemaining && top < area.Bottom && result.Count < OfficeTextLayoutEngine.MaximumLayoutLines) {
            token.ThrowIfCancellationRequested(); accountInspection();
            OfficeRichTextParagraph source = _paragraphs[_paragraph];
            if (_offset == 0 && spacedParagraph != _paragraph) {
                top += source.Margins.Top; spacedParagraph = _paragraph;
            }
            if (top >= area.Bottom) break;
            // Start with the masks at the row's top. Only fonts actually
            // selected into a free interval can enlarge its vertical extent.
            double extent = 0;
            int paragraph = _paragraph, offset = _offset, position = CharacterPosition;
            var row = new List<(OfficeTextFlowRegion Region, IntervalLine Line)>();
            double baseline = 0, below = 0, advanceBelow = 0;
            bool retry;
            do {
                _paragraph = paragraph; _offset = offset; CharacterPosition = position;
                row.Clear(); baseline = below = advanceBelow = 0; retry = false;
                foreach (OfficeTextFlowRegion interval in OfficeTextFlowRegions.Row(area, top, extent, exclusions, accountInspection, token)) {
                    if (result.Count + row.Count >= OfficeTextLayoutEngine.MaximumLayoutLines) break;
                    IntervalLine? line = MeasureInterval(interval.Width, area.X - interval.X, metrics, token);
                    if (line == null) continue;
                    baseline = Math.Max(baseline, line.Baseline);
                    below = Math.Max(below, line.ContentHeight - line.Baseline);
                    advanceBelow = Math.Max(advanceBelow, line.Advance - line.Baseline);
                    if (baseline + below > extent + .000001D) {
                        // A larger font in a later interval can activate a mask
                        // missed by the first fragment. Refit the whole row.
                        extent = baseline + below; retry = true; break;
                    }
                    row.Add((interval, line));
                    ConsumeInterval(line);
                    if (line.EndsRow) break;
                }
            } while (retry);
            if (row.Count == 0 || top + baseline + below > area.Bottom + .000001D) {
                _paragraph = paragraph; _offset = offset; CharacterPosition = position;
                if (row.Count != 0) break;
                double next = area.Bottom;
                foreach (OfficeTextFlowRegion mask in exclusions) {
                    token.ThrowIfCancellationRequested(); accountInspection();
                    if (mask.Width <= 0 || mask.Height <= 0 || mask.Right <= area.X || mask.X >= area.Right) continue;
                    if (mask.Y > top) next = Math.Min(next, mask.Y);
                    if (mask.Bottom > top) next = Math.Min(next, mask.Bottom);
                }
                top = next;
                continue;
            }
            foreach (var fragment in row) result.Add(new OfficeRichTextFlowFragment(fragment.Line.Paragraph,
                new OfficeTextFlowRegion(fragment.Region.X, top + baseline - fragment.Line.Baseline,
                    fragment.Region.Width, fragment.Line.ContentHeight)));
            top += baseline + advanceBelow;
            if (_paragraph != paragraph) top += source.Margins.Bottom;
        }
        return result;
    }

    private IntervalLine? MeasureInterval(double width, double tabOriginAdjustment, OfficeDrawingTextMetrics metrics, CancellationToken token) {
        OfficeRichTextParagraph source = _paragraphs[_paragraph];
        int remaining = TextLength(source) - _offset;
        OfficeRichTextParagraph tail;
        OfficeRichTextBlockLayout layout;
        int lookahead = Math.Min(remaining, 128), length;
        do {
            length = IntervalLookahead(source, _offset, lookahead, remaining, token);
            // Complete words and tab fields retain the breaker's lookahead
            // context. Grow only until a genuine first-line boundary exists;
            // future lines are not laid out afresh for every visible fragment.
            tail = Slice(source, _offset, length, true, verticalMargins: false, tabOriginAdjustment, keepTypingStyles: true);
            _accountMeasurement(length);
            layout = OfficeDrawingTextLayout.CreateFlowParagraphs(new[] { tail }, width,
                double.MaxValue, metrics.MeasureText, metrics, false, token);
            if (length == remaining || layout.Lines.Count > 1 && layout.Lines[0].SourceTextEnd < length) break;
            lookahead = Math.Min(remaining, Math.Max(length + 1, length * 2));
        } while (true);
        if (layout.Lines.Count == 0) return null;
        OfficeRichTextLine line = layout.Lines[0];
        var fitting = line.FittingBounds ?? (line.OffsetX, line.OffsetX + line.Width);
        if (fitting.Item1 < -.000001D || fitting.Item2 > width - tail.Margins.Right + .000001D) return null;
        int consumed = Math.Min(remaining, line.SourceTextEnd);
        bool complete = consumed == remaining && line.CompletesSource;
        if (consumed == 0 && !complete) return null;
        bool emptyLine = true;
        foreach (OfficeRichTextSegment segment in line.Segments) emptyLine &= segment.Text.Length == 0;
        var selected = Slice(source, _offset, consumed, complete, verticalMargins: false, tabOriginAdjustment, keepTypingStyles: emptyLine);
        selected.NativeFlowPlan = new OfficeRichTextFlowPlan(width - source.Margins.Horizontal,
            new[] { (consumed, line.FittingBounds.HasValue ? (double?)(line.FittingBounds.Value.Right - line.FittingBounds.Value.Left) : null,
                line.EndsWithGeneratedHyphen) }, line.NativeTabs);
        double advance = OfficeTextBlockRenderer.ResolveRichTextRenderLineHeight(line, layout.LineHeight);
        double baseline = OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, 0, advance, true);
        bool hardBreak = false;
        if (consumed > 0) {
            int last = _offset + consumed - 1;
            foreach (OfficeRichTextRun run in source.Runs) {
                if (last < run.Text.Length) { hardBreak = run.Text[last] == '\r' || run.Text[last] == '\n'; break; }
                last -= run.Text.Length;
            }
        }
        return new IntervalLine(selected, consumed, complete, complete || hardBreak, advance,
            line.ContentHeight ?? advance, baseline);
    }

    // Stop at a complete token. A tab's alignment can inspect its entire field,
    // so whitespace inside that field cannot end the measurement window.
    private static int IntervalLookahead(OfficeRichTextParagraph source, int start, int target, int remaining, CancellationToken token) {
        if (target >= remaining) return remaining;
        int position = 0, relative = 0;
        bool tabbed = false;
        char previous = '\0';
        foreach (OfficeRichTextRun run in source.Runs) {
            token.ThrowIfCancellationRequested();
            int left = Math.Max(0, start - position);
            for (int i = left; i < run.Text.Length; i++) {
                token.ThrowIfCancellationRequested();
                char c = run.Text[i];
                if (relative >= target && (c == '\t' || c == '\r' || c == '\n' || !tabbed && char.IsWhiteSpace(c)))
                    return c == '\n' && previous == '\r' ? relative - 1 : relative;
                if (c == '\r' || c == '\n') tabbed = false;
                else if (c == '\t') tabbed = true;
                previous = c; relative++;
            }
            position += run.Text.Length;
        }
        return remaining;
    }

    private void ConsumeInterval(IntervalLine line) {
        CharacterPosition = checked(CharacterPosition + line.Consumed); _offset += line.Consumed;
        if (!line.Complete) return;
        _paragraph++; _offset = 0;
        if (HasRemaining) CharacterPosition++;
    }

    private sealed record IntervalLine(OfficeRichTextParagraph Paragraph, int Consumed, bool Complete,
        bool EndsRow, double Advance, double ContentHeight, double Baseline);
}

/// <summary>One immutable line fragment and its placement in the original frame.</summary>
internal sealed record OfficeRichTextFlowFragment(OfficeRichTextParagraph Paragraph, OfficeTextFlowRegion Region);
