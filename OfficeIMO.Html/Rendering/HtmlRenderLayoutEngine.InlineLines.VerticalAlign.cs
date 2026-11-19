namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Line box geometry from CSS 2.1 section 10.8 for a line holding boxes that are not
    /// baseline-aligned. Baseline-relative boxes, including shifted ones, set the line's
    /// ascent and descent around the root strut; top- and bottom-aligned boxes are placed
    /// against the line edges and extend the line away from that edge when they are taller.
    /// </summary>
    private sealed class AlignedInlineLine {
        internal AlignedInlineLine(double height, double baseline, bool normalMetrics, Dictionary<HtmlInlineRun, double> atomicTops,
            Dictionary<object, (double Ascent, double Descent)> lineRelativeGroups) {
            Height = height;
            Baseline = baseline;
            NormalMetrics = normalMetrics;
            AtomicTops = atomicTops;
            LineRelativeGroups = lineRelativeGroups;
        }

        internal double Height { get; }
        internal double Baseline { get; }
        internal bool NormalMetrics { get; }
        internal Dictionary<HtmlInlineRun, double> AtomicTops { get; }
        // Ascent and descent of each top- or bottom-aligned subtree around its own baseline.
        internal Dictionary<object, (double Ascent, double Descent)> LineRelativeGroups { get; }
    }

    private AlignedInlineLine? ResolveAlignedInlineLine(InlineLine line, HtmlRenderBoxStyle paragraphStyle) {
        if (!RequiresAlignedInlineLine(line)) return null;
        bool normalMetrics = paragraphStyle.Font.Size > 0D && paragraphStyle.LineHeight >= paragraphStyle.Font.Size;
        InlineTextLineMetrics strut = ResolveInlineSharedTextMetrics(string.Empty, paragraphStyle, true, normalMetrics);
        double ascent = strut.Baseline;
        double descent = strut.Height - strut.Baseline;
        double topHeight = 0D, bottomHeight = 0D;
        var groups = new Dictionary<object, (double Ascent, double Descent)>();
        var topGroups = new HashSet<object>();
        var atomicTops = new Dictionary<HtmlInlineRun, double>();
        var lineRelativeAtomics = new List<HtmlInlineRun>();
        var groupedAtomics = new List<(HtmlInlineRun Run, object Group, double Top)>();

        foreach (InlineSegment segment in line.Segments) {
            HtmlInlineRun run = segment.Run;
            foreach (HtmlRenderBoxStyle strutStyle in run.InlineStrutStyles) AddInlineBox(string.Empty, strutStyle);
            if (!InlineLine.ContributesFlowContent(run)) continue;
            HtmlRenderFlowBlock? atomic = run.AtomicBlock;
            if (atomic != null) {
                if (atomicTops.ContainsKey(run) || lineRelativeAtomics.Contains(run) || groupedAtomics.Any(item => item.Run == run)) continue;
                string alignment = run.Style.TableVerticalAlignment;
                if (alignment is "top" or "bottom") {
                    lineRelativeAtomics.Add(run);
                    if (alignment == "top") topHeight = Math.Max(topHeight, atomic.Height);
                    else bottomHeight = Math.Max(bottomHeight, atomic.Height);
                    continue;
                }
                double atomicBaseline = run.AtomicBaseline ?? atomic.Height;
                string groupAlign = AtomicGroupAlign(run.Style);
                if (groupAlign.Length > 0) {
                    // Inside a top or bottom inline the box keeps its place relative to that group's baseline.
                    object key = LineRelativeGroupKey(run.Style);
                    double groupTop = ResolveAtomicVerticalShift(run, paragraphStyle, LineRelativeGroupOffset(run.Style),
                        atomic.Height, atomicBaseline) - atomicBaseline;
                    if (groupAlign == "top") topGroups.Add(key);
                    groups.TryGetValue(key, out (double Ascent, double Descent) extent);
                    groups[key] = (Math.Max(extent.Ascent, -groupTop), Math.Max(extent.Descent, atomic.Height + groupTop));
                    groupedAtomics.Add((run, key, groupTop));
                    continue;
                }
                double top = ResolveAtomicVerticalShift(run, paragraphStyle, paragraphStyle.BaselineOffset, atomic.Height, atomicBaseline)
                    - atomicBaseline;
                atomicTops[run] = top;
                ascent = Math.Max(ascent, -top);
                descent = Math.Max(descent, atomic.Height + top);
                continue;
            }
            if (segment.Text.Length == 0) continue;
            AddInlineBox(segment.Text, run.Style);
        }
        // Empty inline elements still contribute their line-height, as CSS 2.1 section 10.8 requires.
        foreach (HtmlRenderBoxStyle strutStyle in line.EmptyInlineStrutStyles) AddInlineBox(string.Empty, strutStyle);

        foreach (KeyValuePair<object, (double Ascent, double Descent)> group in groups) {
            double groupHeight = group.Value.Ascent + group.Value.Descent;
            if (IsTopGroup(group.Key)) topHeight = Math.Max(topHeight, groupHeight);
            else bottomHeight = Math.Max(bottomHeight, groupHeight);
        }
        if (topHeight > ascent + descent) descent = topHeight - ascent;
        if (bottomHeight > ascent + descent) ascent = bottomHeight - descent;
        double height = Math.Max(0.01D, ascent + descent);
        foreach (HtmlInlineRun run in lineRelativeAtomics) {
            // Stored relative to the baseline like the baseline-relative boxes.
            atomicTops[run] = run.Style.TableVerticalAlignment == "top"
                ? -ascent
                : descent - run.AtomicBlock!.Height;
        }
        foreach ((HtmlInlineRun run, object key, double groupTop) in groupedAtomics) {
            double groupBaseline = topGroups.Contains(key) ? groups[key].Ascent : height - groups[key].Descent;
            atomicTops[run] = groupBaseline - ascent + groupTop;
        }
        return new AlignedInlineLine(height, ascent, normalMetrics, atomicTops, groups);

        bool IsTopGroup(object key) => topGroups.Contains(key);

        void AddInlineBox(string text, HtmlRenderBoxStyle style) {
            InlineTextLineMetrics metrics = ResolveInlineSharedTextMetrics(text, style, true, normalMetrics);
            if (style.LineRelativeAlign.Length > 0) {
                // The subtree keeps its baselines: each box extends the group around the group's baseline.
                object key = LineRelativeGroupKey(style);
                double relative = style.BaselineOffset - LineRelativeGroupOffset(style);
                if (style.LineRelativeAlign == "top") topGroups.Add(key);
                groups.TryGetValue(key, out (double Ascent, double Descent) extent);
                groups[key] = (Math.Max(extent.Ascent, metrics.Baseline - relative),
                    Math.Max(extent.Descent, metrics.Height - metrics.Baseline + relative));
            } else {
                double shift = style.BaselineOffset - paragraphStyle.BaselineOffset;
                ascent = Math.Max(ascent, metrics.Baseline - shift);
                descent = Math.Max(descent, metrics.Height - metrics.Baseline + shift);
            }
        }
    }

    private static bool RequiresAlignedInlineLine(InlineLine line) {
        foreach (InlineSegment segment in line.Segments) {
            HtmlInlineRun run = segment.Run;
            if (!InlineLine.ContributesFlowContent(run)) continue;
            if (run.AtomicBlock != null) {
                if (run.Style.TableVerticalAlignment != "baseline" || AtomicGroupAlign(run.Style).Length > 0) return true;
            } else if (segment.Text.Length > 0 && run.Style.LineRelativeAlign.Length > 0) {
                return true;
            }
        }
        return false;
    }

    // Downward shift of an atomic inline box from baseline alignment. middle, text-top and
    // text-bottom use the box's own height against the parent's font; lengths, percentages,
    // sub and super were resolved into BaselineOffset with the element's style.
    // baseOffset is the BaselineOffset of the baseline the box is placed against: the paragraph's, or that of
    // the top or bottom inline whose group it belongs to.
    private static double ResolveAtomicVerticalShift(HtmlInlineRun run, HtmlRenderBoxStyle paragraphStyle, double baseOffset,
        double height, double atomicBaseline) {
        HtmlRenderBoxStyle style = run.Style;
        double inherited = style.BaselineOffset - style.VerticalAlignShift - baseOffset;
        double parentFontSize = double.IsNaN(style.VerticalAlignParentFontSize) ? paragraphStyle.Font.Size : style.VerticalAlignParentFontSize;
        return style.TableVerticalAlignment switch {
            "middle" => inherited + atomicBaseline - height / 2D - parentFontSize * 0.25D,
            "text-top" => inherited + atomicBaseline - parentFontSize * 0.8D,
            "text-bottom" => inherited + parentFontSize * 0.2D - (height - atomicBaseline),
            _ => style.BaselineOffset - baseOffset
        };
    }

    // Replaced inline boxes are display inline and inherit the group directly; inline-blocks record it separately.
    private static string AtomicGroupAlign(HtmlRenderBoxStyle style) =>
        style.TableVerticalAlignment is "top" or "bottom" ? string.Empty
            : style.AtomicLineRelativeAlign.Length > 0 ? style.AtomicLineRelativeAlign : style.LineRelativeAlign;

    /// <summary>The paint baseline of a text segment on an aligned line, before its own BaselineOffset is applied.</summary>
    private double ResolveAlignedTextBaseline(AlignedInlineLine aligned, InlineSegment segment) {
        HtmlRenderBoxStyle style = segment.Run.Style;
        if (style.LineRelativeAlign.Length == 0) return aligned.Baseline;
        if (!aligned.LineRelativeGroups.TryGetValue(LineRelativeGroupKey(style), out (double Ascent, double Descent) group)) {
            InlineTextLineMetrics metrics = ResolveInlineSharedTextMetrics(segment.Text, style, true, aligned.NormalMetrics);
            group = (metrics.Baseline, metrics.Height - metrics.Baseline);
        }
        double groupBaseline = style.LineRelativeAlign == "top" ? group.Ascent : aligned.Height - group.Descent;
        // Paint adds the segment's own BaselineOffset back, which places it relative to the group's baseline.
        return groupBaseline - LineRelativeGroupOffset(style);
    }

    // Styles without a recorded group, such as synthetic ones, form a group of their own.
    private static object LineRelativeGroupKey(HtmlRenderBoxStyle style) => style.LineRelativeAlignRoot ?? style;

    private static double LineRelativeGroupOffset(HtmlRenderBoxStyle style) =>
        style.LineRelativeAlignRoot != null ? style.LineRelativeAlignOffset : style.BaselineOffset;
}
