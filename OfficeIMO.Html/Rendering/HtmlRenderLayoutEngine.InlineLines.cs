using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private sealed class InlineLine {
        private int _flowContentCount;
        private double _availableWidth;
        private double _indent;
        internal double IndentOffset { get; private set; }

        internal List<InlineSegment> Segments { get; } = new List<InlineSegment>();
        internal double Width { get; private set; }
        internal bool HasFlowContent => _flowContentCount > 0;
        internal bool HasExplicitPlacement { get; private set; }
        internal double X { get; private set; }
        internal double Y { get; private set; }
        internal double AvailableWidth => Math.Max(0.01D, _availableWidth - _indent);
        internal bool EndsWithHyphenation { get; set; }

        internal void Place(double x, double y, double availableWidth) {
            HasExplicitPlacement = true;
            X = Math.Max(0D, x);
            Y = Math.Max(0D, y);
            _availableWidth = Math.Max(0.01D, availableWidth);
        }

        internal void Indent(double indent, bool rightToLeft) {
            _indent = indent;
            IndentOffset = rightToLeft ? 0D : indent;
        }

        internal double ResolveAvailableWidth(double width) => HasExplicitPlacement
            ? AvailableWidth : Math.Max(0.01D, width - _indent);

        internal double ResolveAlignmentOffset(OfficeTextAlignment alignment, double width) {
            if (_indent == 0D) return ResolveLineOffset(alignment, ResolveAvailableWidth(width), Width);
            // An indent can exceed the line width. Preserve the signed space for
            // alignment so RTL text overflows towards inline-end, not the right.
            double remaining = (HasExplicitPlacement ? _availableWidth : width) - _indent - Width;
            if (alignment == OfficeTextAlignment.Right) return remaining;
            if (alignment == OfficeTextAlignment.Center) return remaining / 2D;
            return 0D;
        }

        internal void Add(InlineSegment segment) {
            Segments.Add(segment);
            Width += segment.Width;
            if (segment.Run.RunningStringElement == null
                && segment.Run.RunningElementAssignment == null
                && !segment.Run.IsBookmarkMarker) _flowContentCount++;
        }

        internal void RemoveAt(int index) {
            if (Segments[index].Run.RunningStringElement == null
                && Segments[index].Run.RunningElementAssignment == null
                && !Segments[index].Run.IsBookmarkMarker) _flowContentCount--;
            Width -= Segments[index].Width;
            Segments.RemoveAt(index);
        }

        internal void SetSegmentWidth(int index, double width) {
            double normalized = Math.Max(0D, width);
            Width += normalized - Segments[index].Width;
            Segments[index].SetWidth(normalized);
        }

        internal double ResolveLineHeight(double fallback) {
            if (!HasFlowContent) return 0D;
            double height = fallback;
            for (int i = 0; i < Segments.Count; i++) {
                height = Math.Max(height, Segments[i].Run.AtomicBlock?.Height ?? Segments[i].Run.Style.LineHeight);
            }
            if (!HasReplacedImage) {
                if (Segments.Count > 0 && HasMixedTextSizes(Segments[0].Run.Style)) {
                    double textBaseline = 0D;
                    double textDescent = 0D;
                    foreach (InlineSegment segment in Segments) {
                        if (segment.Run.AtomicBlock == null) {
                            textBaseline = Math.Max(textBaseline, segment.Run.Style.Font.Size);
                            textDescent = Math.Max(textDescent, segment.Run.Style.LineHeight - segment.Run.Style.Font.Size);
                        }
                    }
                    height = Math.Max(height, textBaseline + textDescent);
                }
                return Math.Max(0.01D, height);
            }

            double ascent = 0D;
            double descent = 0D;
            for (int i = 0; i < Segments.Count; i++) {
                HtmlInlineRun run = Segments[i].Run;
                if (run.AtomicBlock != null) {
                    double atomicBaseline = Math.Min(run.AtomicBlock.Height, Math.Max(0D, run.AtomicBaseline ?? run.AtomicBlock.Height));
                    ascent = Math.Max(ascent, atomicBaseline);
                    descent = Math.Max(descent, run.AtomicBlock.Height - atomicBaseline);
                } else {
                    ascent = Math.Max(ascent, run.Style.Font.Size);
                    descent = Math.Max(descent, Math.Max(0D, run.Style.LineHeight - run.Style.Font.Size));
                }
            }
            return Math.Max(0.01D, Math.Max(height, ascent + descent));
        }

        internal bool HasReplacedImage => Segments.Any(segment => segment.Run.IsReplacedImage);

        internal bool HasMixedTextSizes(HtmlRenderBoxStyle paragraphStyle) =>
            Segments.Any(segment => segment.Run.AtomicBlock == null
                && Math.Abs(segment.Run.Style.Font.Size - paragraphStyle.Font.Size) > 0.000001D);

        internal double ResolveBaseline(HtmlRenderBoxStyle paragraphStyle) {
            if (!HasReplacedImage) {
                // Positioned text in the shared drawing model paints its baseline
                // at Y + the source font size. Keep mixed-size text on that same
                // baseline, including the containing paragraph's line strut.
                double baseline = paragraphStyle.Font.Size;
                foreach (InlineSegment segment in Segments) {
                    if (segment.Run.AtomicBlock == null) baseline = Math.Max(baseline, segment.Run.Style.Font.Size);
                }
                return baseline;
            }
            double ascent = paragraphStyle.Font.Size;
            for (int i = 0; i < Segments.Count; i++) {
                HtmlInlineRun run = Segments[i].Run;
                ascent = Math.Max(ascent, run.AtomicBlock == null
                    ? run.Style.Font.Size
                    : Math.Min(run.AtomicBlock.Height, Math.Max(0D, run.AtomicBaseline ?? run.AtomicBlock.Height)));
            }
            return ascent;
        }
    }

    private static OfficeFontInfo GetEffectiveTextFont(HtmlRenderBoxStyle style) =>
        Math.Abs(style.BaselineScale - 1D) < 0.000001D
            ? style.Font
            : style.Font.WithSize(style.Font.Size * style.BaselineScale);
}
