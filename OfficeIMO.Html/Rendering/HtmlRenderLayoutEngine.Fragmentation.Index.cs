using System.Runtime.CompilerServices;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Fewer children are cheaper to test one by one than to index.
    private const int IndexedSliceChildCount = 32;
    private static readonly ConditionalWeakTable<HtmlRenderVisual, SliceIndex> SliceIndexes = new();

    /// <summary>The bottom a visual reaches for page slicing, including children that continue past its box.</summary>
    private static double SliceBottom(HtmlRenderVisual visual) {
        double visualBottom = visual.LayoutY + visual.LayoutHeight;
        // Paint-neutral wrappers can retain the CSS box height while
        // their visible children continue across printed pages.
        if (visual is HtmlRenderSemanticGroup or HtmlRenderLayoutRegion or HtmlRenderClipGroup { ClipVertical: false }) {
            visualBottom = Math.Max(visualBottom, ChildScrollBottom(visual)!.Value);
        } else if (visual is HtmlRenderEffectGroup effectOverflow) {
            visualBottom = Math.Max(visualBottom, Math.Max(0D, ScrollBottom(effectOverflow)));
        }
        return visualBottom;
    }

    /// <summary>
    /// The children a page slice must examine, in source order. A container with many children, such as a table
    /// that spans hundreds of pages, is indexed once by vertical extent so each page visits only the children that
    /// can overlap it. Every child that can add a fragment is returned, so the slice itself is unchanged.
    /// </summary>
    private IEnumerable<HtmlRenderVisual> SliceCandidates(IEnumerable<HtmlRenderVisual> sourceVisuals,
        HtmlRenderVisual? owner, double start, double end) {
        // Page floats can anchor inside any child, so their pages examine every child.
        if (owner == null || _pageFloatEntries.Count > 0 || _columnEdgeFloatEntries.Count > 0
            || sourceVisuals is not IReadOnlyList<HtmlRenderVisual> { Count: >= IndexedSliceChildCount } list) {
            return sourceVisuals;
        }
        SliceIndex index = SliceIndexes.GetValue(owner, _ => new SliceIndex(list));
        // The owner's children are immutable; a different list is not the one indexed.
        return ReferenceEquals(index.Source, list) ? index.Select(start, end) : sourceVisuals;
    }

    private sealed class SliceIndex {
        private readonly int[] _byTop;
        private readonly double[] _sortedTops;
        private readonly double[] _bottoms;
        private readonly int[] _always;
        private readonly int[] _byBoundary;
        private readonly double[] _sortedBoundaries;
        private readonly double _maximumShortHeight;

        internal SliceIndex(IReadOnlyList<HtmlRenderVisual> source) {
            Source = source;
            int count = source.Count;
            var tops = new double[count];
            _bottoms = new double[count];
            var heights = new double[count];
            var boundaries = new List<(double Y, int Index)>();
            for (int index = 0; index < count; index++) {
                HtmlRenderVisual visual = source[index];
                tops[index] = visual.LayoutY;
                _bottoms[index] = SliceBottom(visual);
                heights[index] = Math.Max(0D, _bottoms[index] - tops[index]);
                // A resumed fragment keeps a collapsed rule that starts at its top, even without overlap.
                if (visual is HtmlRenderShape { CollapsedBorderBoundaryY: double boundary }) boundaries.Add((boundary, index));
            }
            double[] sortedHeights = (double[])heights.Clone();
            Array.Sort(sortedHeights);
            // Children far taller than the rows, such as a rule or background spanning the whole table, are kept
            // for every page. The rest are found by their top within one short height of the page window.
            double tallThreshold = Math.Max(4D * sortedHeights[(int)((count - 1) * 0.95D)], 1D);
            var always = new List<int>();
            var shortChildren = new List<int>(count);
            for (int index = 0; index < count; index++) {
                if (heights[index] > tallThreshold) {
                    always.Add(index);
                } else {
                    shortChildren.Add(index);
                    _maximumShortHeight = Math.Max(_maximumShortHeight, heights[index]);
                }
            }
            _always = always.ToArray();
            _byTop = shortChildren.ToArray();
            _sortedTops = new double[_byTop.Length];
            for (int index = 0; index < _byTop.Length; index++) _sortedTops[index] = tops[_byTop[index]];
            Array.Sort(_sortedTops, _byTop);
            _byBoundary = new int[boundaries.Count];
            _sortedBoundaries = new double[boundaries.Count];
            for (int index = 0; index < boundaries.Count; index++) {
                _sortedBoundaries[index] = boundaries[index].Y;
                _byBoundary[index] = boundaries[index].Index;
            }
            Array.Sort(_sortedBoundaries, _byBoundary);
        }

        internal IReadOnlyList<HtmlRenderVisual> Source { get; }

        internal IReadOnlyList<HtmlRenderVisual> Select(double start, double end) {
            // A short child overlaps [start, end] only when its top lies after start minus its height and no later
            // than end. Flow anchors are found the same way. These bounds are wider than the 0.0001 slicing
            // tolerance; SliceVisuals applies its exact tests to every child returned here.
            int first = LowerBound(_sortedTops, start - _maximumShortHeight - 0.001D);
            int last = LowerBound(_sortedTops, end + 0.001D);
            var selected = new List<int>(_always.Length + Math.Max(0, last - first));
            selected.AddRange(_always);
            for (int position = first; position < last; position++) {
                int index = _byTop[position];
                if (_bottoms[index] >= start - 0.001D) selected.Add(index);
            }
            for (int position = LowerBound(_sortedBoundaries, start - 0.001D);
                position < _sortedBoundaries.Length && _sortedBoundaries[position] <= start + 0.001D; position++) {
                selected.Add(_byBoundary[position]);
            }
            selected.Sort();
            var visuals = new List<HtmlRenderVisual>(selected.Count);
            for (int position = 0; position < selected.Count; position++) {
                // A rule found both by overlap and by its boundary is examined once.
                if (position > 0 && selected[position] == selected[position - 1]) continue;
                visuals.Add(Source[selected[position]]);
            }
            return visuals;
        }

        private static int LowerBound(double[] values, double value) {
            int low = 0;
            int high = values.Length;
            while (low < high) {
                int middle = low + ((high - low) >> 1);
                if (values[middle] < value) low = middle + 1;
                else high = middle;
            }
            return low;
        }
    }
}
