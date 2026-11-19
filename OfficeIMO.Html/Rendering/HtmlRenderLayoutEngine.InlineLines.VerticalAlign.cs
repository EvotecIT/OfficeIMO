namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Line box geometry from CSS 2.1 section 10.8 for a line holding boxes that are not
    /// baseline-aligned. Baseline-relative boxes, including shifted ones, set the line's
    /// ascent and descent around the root strut; top- and bottom-aligned boxes are placed
    /// against the line edges and extend the line away from that edge when they are taller.
    /// </summary>
    private sealed class AlignedInlineLine {
        internal AlignedInlineLine(double height, double baseline, bool normalMetrics, Dictionary<HtmlInlineRun, double> atomicTops) {
            Height = height;
            Baseline = baseline;
            NormalMetrics = normalMetrics;
            AtomicTops = atomicTops;
        }

        internal double Height { get; }
        internal double Baseline { get; }
        internal bool NormalMetrics { get; }
        internal Dictionary<HtmlInlineRun, double> AtomicTops { get; }
    }

    private AlignedInlineLine? ResolveAlignedInlineLine(InlineLine line, HtmlRenderBoxStyle paragraphStyle) {
        if (!RequiresAlignedInlineLine(line)) return null;
        bool normalMetrics = paragraphStyle.Font.Size > 0D && paragraphStyle.LineHeight >= paragraphStyle.Font.Size;
        InlineTextLineMetrics strut = ResolveInlineSharedTextMetrics(string.Empty, paragraphStyle, true, normalMetrics);
        double ascent = strut.Baseline;
        double descent = strut.Height - strut.Baseline;
        double topHeight = 0D, bottomHeight = 0D;
        var atomicTops = new Dictionary<HtmlInlineRun, double>();
        var lineRelativeAtomics = new List<HtmlInlineRun>();

        foreach (InlineSegment segment in line.Segments) {
            HtmlInlineRun run = segment.Run;
            if (!InlineLine.ContributesFlowContent(run)) continue;
            HtmlRenderFlowBlock? atomic = run.AtomicBlock;
            if (atomic != null) {
                if (atomicTops.ContainsKey(run) || lineRelativeAtomics.Contains(run)) continue;
                string alignment = run.Style.TableVerticalAlignment;
                if (alignment is "top" or "bottom") {
                    lineRelativeAtomics.Add(run);
                    if (alignment == "top") topHeight = Math.Max(topHeight, atomic.Height);
                    else bottomHeight = Math.Max(bottomHeight, atomic.Height);
                    continue;
                }
                double atomicBaseline = run.AtomicBaseline ?? atomic.Height;
                double top = ResolveAtomicVerticalShift(run, paragraphStyle, atomic.Height, atomicBaseline) - atomicBaseline;
                atomicTops[run] = top;
                ascent = Math.Max(ascent, -top);
                descent = Math.Max(descent, atomic.Height + top);
                continue;
            }
            if (segment.Text.Length == 0) continue;
            HtmlRenderBoxStyle style = run.Style;
            InlineTextLineMetrics metrics = ResolveInlineSharedTextMetrics(segment.Text, style, true, normalMetrics);
            if (style.LineRelativeAlign == "top") {
                topHeight = Math.Max(topHeight, metrics.Height);
            } else if (style.LineRelativeAlign == "bottom") {
                bottomHeight = Math.Max(bottomHeight, metrics.Height);
            } else {
                double shift = style.BaselineOffset - paragraphStyle.BaselineOffset;
                ascent = Math.Max(ascent, metrics.Baseline - shift);
                descent = Math.Max(descent, metrics.Height - metrics.Baseline + shift);
            }
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
        return new AlignedInlineLine(height, ascent, normalMetrics, atomicTops);
    }

    private static bool RequiresAlignedInlineLine(InlineLine line) {
        foreach (InlineSegment segment in line.Segments) {
            HtmlInlineRun run = segment.Run;
            if (!InlineLine.ContributesFlowContent(run)) continue;
            if (run.AtomicBlock != null) {
                if (run.Style.TableVerticalAlignment != "baseline") return true;
            } else if (segment.Text.Length > 0 && run.Style.LineRelativeAlign.Length > 0) {
                return true;
            }
        }
        return false;
    }

    // Downward shift of an atomic inline box from baseline alignment. middle, text-top and
    // text-bottom use the box's own height against the parent's font; lengths, percentages,
    // sub and super were resolved into BaselineOffset with the element's style.
    private static double ResolveAtomicVerticalShift(HtmlInlineRun run, HtmlRenderBoxStyle paragraphStyle,
        double height, double atomicBaseline) {
        HtmlRenderBoxStyle style = run.Style;
        double inherited = style.BaselineOffset - style.VerticalAlignShift - paragraphStyle.BaselineOffset;
        double parentFontSize = style.VerticalAlignParentFontSize > 0D ? style.VerticalAlignParentFontSize : paragraphStyle.Font.Size;
        return style.TableVerticalAlignment switch {
            "middle" => inherited + atomicBaseline - height / 2D - parentFontSize * 0.25D,
            "text-top" => inherited + atomicBaseline - parentFontSize * 0.8D,
            "text-bottom" => inherited + parentFontSize * 0.2D - (height - atomicBaseline),
            _ => style.BaselineOffset - paragraphStyle.BaselineOffset
        };
    }

    /// <summary>The paint baseline of a text segment on an aligned line, before its own BaselineOffset is applied.</summary>
    private double ResolveAlignedTextBaseline(AlignedInlineLine aligned, InlineSegment segment) {
        HtmlRenderBoxStyle style = segment.Run.Style;
        if (style.LineRelativeAlign.Length == 0) return aligned.Baseline;
        InlineTextLineMetrics metrics = ResolveInlineSharedTextMetrics(segment.Text, style, true, aligned.NormalMetrics);
        double baseline = style.LineRelativeAlign == "top"
            ? metrics.Baseline
            : aligned.Height - (metrics.Height - metrics.Baseline);
        // Line-relative boxes ignore inherited baseline shifts, which paint would add back.
        return baseline - style.BaselineOffset;
    }
}
