using System.Collections.ObjectModel;
using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyDictionary<IElement, ColumnSourceFragment>? _columnSourceSelection;

    /// <summary>
    /// Splits source flow around spanners without changing the DOM that owns
    /// selector matching, inheritance, generated content and source identity.
    /// </summary>
    private IReadOnlyList<MultiColumnSpanPartition> ResolveMultiColumnSpanPartitions(
        IElement element, double contentWidth, HtmlRenderBoxStyle parentStyle,
        double? columnWidth = null, int depth = 0) {
        double gap = parentStyle.ColumnGapWasSpecified ? parentStyle.ColumnGap : parentStyle.Font.Size;
        int count = ResolveRequestedColumnCount(parentStyle, contentWidth, gap);
        double sourceWidth = columnWidth ?? Math.Max(0.01D,
            (contentWidth - gap * Math.Max(0, count - 1)) / count);
        return SplitColumnSourceNodes(element, sourceWidth, contentWidth, parentStyle, depth);
    }

    private IReadOnlyList<MultiColumnSpanPartition> SplitColumnSourceNodes(
        IElement container, double sourceWidth, double spanningWidth,
        HtmlRenderBoxStyle parentStyle, int depth) {
        EnsureDepth(depth, container);
        var partitions = new List<MultiColumnSpanPartition>();
        var nodes = new List<INode>();
        var selection = new Dictionary<IElement, ColumnSourceFragment>();
        foreach (INode node in container.ChildNodes) {
            CheckCancellation();
            ChargeLayoutOperation("column source partition discovery");
            if (node is not IElement child || ShouldSkipElement(child) || IsClosedDisclosureChild(node)) {
                nodes.Add(node);
                continue;
            }

            HtmlRenderBoxStyle childStyle = ForwardContainingHeightBasis(child,
                _styleResolver.Resolve(child, sourceWidth, parentStyle), parentStyle);
            if (IsColumnSourceSpanner(child, childStyle)) {
                HtmlRenderBoxStyle spanningStyle = ForwardContainingHeightBasis(child,
                    _styleResolver.Resolve(child, spanningWidth, parentStyle), parentStyle);
                partitions.Add(new MultiColumnSpanPartition(nodes, child, spanningStyle,
                    selection, Array.Empty<ColumnSpannerAncestor>(), parentStyle, depth + 1));
                nodes = new List<INode>();
                selection = new Dictionary<IElement, ColumnSourceFragment>();
                continue;
            }

            if (!CanTraverseColumnSpannerSource(child, childStyle)) {
                nodes.Add(node);
                continue;
            }

            childStyle = ResolveOrdinaryIntrinsicWidths(child, childStyle, sourceWidth, depth + 1);
            double childWidth = sourceWidth;
            if (childStyle.Display is not ("inline" or "contents")) {
                double boxWidth = ResolveBoxWidth(Math.Max(1D,
                    sourceWidth - childStyle.MarginLeft - childStyle.MarginRight), childStyle);
                childWidth = Math.Max(1D, boxWidth - childStyle.HorizontalInsets);
            }
            IReadOnlyList<MultiColumnSpanPartition> descendants = SplitColumnSourceNodes(
                child, childWidth, spanningWidth, childStyle, depth + 1);
            if (descendants.Count == 0) {
                nodes.Add(node);
                continue;
            }

            var ancestor = new ColumnSpannerAncestor(child, childStyle, parentStyle, sourceWidth, depth + 1);
            for (int index = 0; index < descendants.Count; index++) {
                CheckCancellation();
                ChargeLayoutOperation("column source ancestor projection");
                MultiColumnSpanPartition descendant = descendants[index];
                // Empty edge fragments still own initial/final wrapper edges,
                // generated content, margins and source navigation anchors.
                nodes.Add(child);
                selection.Add(child, new ColumnSourceFragment(child, descendant.Nodes,
                    index == 0, index == descendants.Count - 1, ancestor));
                foreach (KeyValuePair<IElement, ColumnSourceFragment> entry in descendant.Selection) {
                    CheckCancellation();
                    ChargeLayoutOperation("column source selection projection");
                    selection.Add(entry.Key, entry.Value);
                }
                if (descendant.Spanner == null) continue;

                var ancestors = new List<ColumnSpannerAncestor> { ancestor };
                foreach (ColumnSpannerAncestor outer in descendant.SpannerAncestors) {
                    CheckCancellation();
                    ChargeLayoutOperation("column spanner ancestor path");
                    ancestors.Add(outer);
                }
                partitions.Add(new MultiColumnSpanPartition(nodes, descendant.Spanner,
                    descendant.SpannerStyle, selection, ancestors,
                    descendant.SpannerParentStyle, descendant.SpannerDepth));
                nodes = new List<INode>();
                selection = new Dictionary<IElement, ColumnSourceFragment>();
            }
        }

        if (partitions.Count == 0) return Array.Empty<MultiColumnSpanPartition>();
        partitions.Add(new MultiColumnSpanPartition(nodes, null, null, selection,
            Array.Empty<ColumnSpannerAncestor>(), parentStyle, depth));
        return partitions.AsReadOnly();
    }

    private static bool IsColumnSourceSpanner(IElement element, HtmlRenderBoxStyle style) =>
        style.Display != "none" && style.ColumnSpan == "all" && style.FloatSide == "none"
        && !ShouldExtractOutOfFlow(style)
        && !HtmlCssRunningElementParser.TryParsePosition(style.Position, out _)
        && HtmlRenderStyleResolver.IsBlockElement(element, style);

    private bool CanTraverseColumnSpannerSource(IElement element, HtmlRenderBoxStyle style) {
        if (style.Display is not ("block" or "inline" or "contents" or "list-item")
            || HtmlCssRunningElementParser.TryParsePosition(style.Position, out _)
            || style.WritingMode != "horizontal-tb") return false;
        // Boxless contents and ordinary non-replaced inlines do not acquire a
        // new block formatting context merely from unused overflow/column values.
        if (style.Display != "contents" && (style.FloatSide != "none" || ShouldExtractOutOfFlow(style)
            || EstablishesTransformedContainingBlock(style)
            || style.Display != "inline" && (HasMultiColumnLayout(style)
                || EstablishesIndependentFloatContext(style)))) return false;
        string tag = element.LocalName;
        return !IsReplacedImageElementTag(tag) && !IsFormControlElement(tag)
            && tag != "math" && tag != "hr";
    }

    /// <summary>Activates only the current immutable source selection for nested ordinary layout.</summary>
    private IDisposable EnterColumnSourcePartition(MultiColumnSpanPartition partition) =>
        new ColumnSourceProjectionScope(this, partition.Selection);

    private IEnumerable<INode> SelectedColumnChildNodes(IElement element) =>
        TryGetColumnSourceFragment(element, out ColumnSourceFragment fragment)
            ? fragment.Nodes : element.ChildNodes;

    private bool TryGetColumnSourceFragment(IElement element, out ColumnSourceFragment fragment) {
        if (_columnSourceSelection != null && _columnSourceSelection.TryGetValue(element, out fragment!)) return true;
        fragment = null!;
        return false;
    }

    private bool IncludeColumnGeneratedContent(IElement element, HtmlPseudoElementKind kind) =>
        !TryGetColumnSourceFragment(element, out ColumnSourceFragment fragment)
        || (kind == HtmlPseudoElementKind.After ? fragment.LastFragment : fragment.FirstFragment);

    /// <summary>Retains authored styles while removing only edges and controls absent from this source fragment.</summary>
    private HtmlRenderBoxStyle ResolveColumnFragmentStyle(IElement element, HtmlRenderBoxStyle style) {
        if (!TryGetColumnSourceFragment(element, out ColumnSourceFragment fragment)) return style;
        HtmlRenderBoxStyle result = style.Clone();
        if (!fragment.FirstFragment) {
            result.BreakBefore = HtmlPageBreakTarget.None;
            result.AvoidBreakBefore = false;
            result.AvoidBreakBeforeScope = HtmlFragmentationScope.None;
            result.StringSet = string.Empty;
        }
        if (!fragment.LastFragment) {
            result.BreakAfter = HtmlPageBreakTarget.None;
            result.AvoidBreakAfter = false;
            result.AvoidBreakAfterScope = HtmlFragmentationScope.None;
        }
        if (result.BoxDecorationBreak == "clone") return result;
        if (!fragment.FirstFragment) {
            result.MarginTop = 0D;
            result.MarginTopAuto = false;
            result.PaddingTop = 0D;
            result.Borders = new HtmlRenderBorderEdges(result.Borders.Top.WithWidth(0D),
                result.Borders.Right, result.Borders.Bottom, result.Borders.Left);
            result.BorderTopLeftRadius = "0";
            result.BorderTopRightRadius = "0";
        }
        if (!fragment.LastFragment) {
            result.MarginBottom = 0D;
            result.MarginBottomAuto = false;
            result.PaddingBottom = 0D;
            result.Borders = new HtmlRenderBorderEdges(result.Borders.Top, result.Borders.Right,
                result.Borders.Bottom.WithWidth(0D), result.Borders.Left);
            result.BorderBottomLeftRadius = "0";
            result.BorderBottomRightRadius = "0";
        }
        return result;
    }

    /// <summary>One immutable view of an original ancestor between successive extracted spanners.</summary>
    private sealed class ColumnSourceFragment {
        internal ColumnSourceFragment(IElement element, IEnumerable<INode> nodes,
            bool firstFragment, bool lastFragment, ColumnSpannerAncestor ancestor) {
            Element = element;
            Nodes = new List<INode>(nodes).AsReadOnly();
            FirstFragment = firstFragment;
            LastFragment = lastFragment;
            Ancestor = ancestor;
        }

        internal IElement Element { get; }
        internal IReadOnlyList<INode> Nodes { get; }
        internal bool FirstFragment { get; }
        internal bool LastFragment { get; }
        internal ColumnSpannerAncestor Ancestor { get; }
    }

    /// <summary>The original ancestor's layout basis, independent of any fragment-edge suppression.</summary>
    private sealed class ColumnSpannerAncestor {
        internal ColumnSpannerAncestor(IElement element, HtmlRenderBoxStyle style,
            HtmlRenderBoxStyle parentStyle, double containingWidth, int depth) {
            Element = element;
            Style = style.Clone();
            ParentStyle = parentStyle.Clone();
            ContainingWidth = containingWidth;
            Depth = depth;
        }

        internal IElement Element { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal HtmlRenderBoxStyle ParentStyle { get; }
        internal double ContainingWidth { get; }
        internal int Depth { get; }
    }

    private sealed class MultiColumnSpanPartition {
        internal MultiColumnSpanPartition(IEnumerable<INode> nodes, IElement? spanner,
            HtmlRenderBoxStyle? spannerStyle, IDictionary<IElement, ColumnSourceFragment> selection,
            IEnumerable<ColumnSpannerAncestor> spannerAncestors, HtmlRenderBoxStyle spannerParentStyle,
            int spannerDepth) {
            Nodes = new List<INode>(nodes).AsReadOnly();
            Spanner = spanner;
            SpannerStyle = spannerStyle;
            Selection = new ReadOnlyDictionary<IElement, ColumnSourceFragment>(
                new Dictionary<IElement, ColumnSourceFragment>(selection));
            SpannerAncestors = new List<ColumnSpannerAncestor>(spannerAncestors).AsReadOnly();
            SpannerParentStyle = spannerParentStyle.Clone();
            SpannerDepth = spannerDepth;
        }

        internal IReadOnlyList<INode> Nodes { get; }
        internal IElement? Spanner { get; }
        internal HtmlRenderBoxStyle? SpannerStyle { get; }
        internal IReadOnlyDictionary<IElement, ColumnSourceFragment> Selection { get; }
        internal IReadOnlyList<ColumnSpannerAncestor> SpannerAncestors { get; }
        internal HtmlRenderBoxStyle SpannerParentStyle { get; }
        internal int SpannerDepth { get; }
    }

    private sealed class ColumnSourceProjectionScope : IDisposable {
        private readonly HtmlRenderLayoutEngine _owner;
        private readonly IReadOnlyDictionary<IElement, ColumnSourceFragment>? _previous;
        private bool _disposed;

        internal ColumnSourceProjectionScope(HtmlRenderLayoutEngine owner,
            IReadOnlyDictionary<IElement, ColumnSourceFragment> selection) {
            _owner = owner;
            _previous = owner._columnSourceSelection;
            owner._columnSourceSelection = selection;
        }

        public void Dispose() {
            if (_disposed) return;
            _owner._columnSourceSelection = _previous;
            _disposed = true;
        }
    }
}
