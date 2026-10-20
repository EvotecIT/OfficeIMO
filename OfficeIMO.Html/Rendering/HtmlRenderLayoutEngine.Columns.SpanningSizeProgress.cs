using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private ColumnSourceFlowContext? _columnSourceFlowContext;

    /// <summary>Owns consumed source-box size across all partitions of one multicol layout attempt.</summary>
    private IDisposable EnterColumnSourceFlow() => new ColumnSourceFlowScope(this);

    /// <summary>
    /// Divides an original box's height constraints between its source pieces.
    /// Repeated measurement of this piece excludes its previous completed size.
    /// </summary>
    private HtmlRenderBoxStyle ResolveColumnSourceFragmentSize(ColumnSourceFragment fragment,
        HtmlRenderBoxStyle style) {
        if (_columnSourceFlowContext == null || style.Display is "inline" or "contents") return style;
        HtmlRenderBoxStyle original = fragment.Ancestor.Style;
        bool hasTarget = original.ExplicitHeight.HasValue || original.AspectRatio.HasValue;
        if (!hasTarget && !original.MinHeight.HasValue && !original.MaxHeight.HasValue) return style;
        CheckCancellation();
        ChargeLayoutOperation("column source box-size allocation");
        double boxWidth = ResolveBoxWidth(Math.Max(1D, fragment.Ancestor.ContainingWidth
            - original.MarginLeft - original.MarginRight), original);
        double consumed = _columnSourceFlowContext.ConsumedBefore(fragment);
        double? remainingTarget = hasTarget
            ? Math.Max(0D, ResolveBoxHeight(0D, boxWidth, original) - consumed) : null;
        double? remainingMinimum = null;
        if (original.MinHeight.HasValue && fragment.LastFragment) {
            HtmlRenderBoxStyle minimum = original.Clone();
            minimum.ExplicitHeight = null;
            minimum.AspectRatio = null;
            minimum.MaxHeight = null;
            remainingMinimum = Math.Max(0D, ResolveBoxHeight(0D, boxWidth, minimum) - consumed);
        }
        double? remainingMaximum = remainingTarget;
        if (original.MaxHeight.HasValue) {
            HtmlRenderBoxStyle maximum = original.Clone();
            maximum.ExplicitHeight = original.MaxHeight;
            maximum.AspectRatio = null;
            maximum.MaxHeight = null;
            double maximumBorderBox = Math.Max(0D, ResolveBoxHeight(0D, boxWidth, maximum) - consumed);
            remainingMaximum = remainingMaximum.HasValue
                ? Math.Min(remainingMaximum.Value, maximumBorderBox) : maximumBorderBox;
        }
        if (!fragment.LastFragment && style.BoxDecorationBreak != "clone" && remainingMaximum.HasValue) {
            // The source-last fragment still owns its physical closing edge.
            // Overflow in an earlier piece must not consume that edge's budget.
            double finalInsets = original.PaddingBottom + original.BorderBottomWidth;
            remainingMaximum = Math.Max(0D, remainingMaximum.Value - finalInsets);
        }

        // Earlier pieces advance only through their used content. The final
        // piece receives any blank tail from an authored height or minimum.
        style.AspectRatio = null;
        style.ExplicitHeight = fragment.LastFragment && remainingTarget.HasValue
            ? ToSourceHeight(remainingTarget.Value) : null;
        style.MinHeight = remainingMinimum.HasValue ? ToSourceHeight(remainingMinimum.Value) : null;
        style.MaxHeight = remainingMaximum.HasValue ? ToSourceHeight(remainingMaximum.Value) : null;
        // Dividing flow allocation does not change the authored containing
        // height against which the original descendants resolve percentages.
        style.HasForwardedContainingHeight = true;
        style.ForwardedContainingHeight = original.DefiniteContentHeight;
        return style;

        double ToSourceHeight(double borderBox) => style.BorderBox
            ? borderBox : Math.Max(0D, borderBox - style.VerticalInsets);
    }

    /// <summary>Records the finalized box advance, replacing an earlier probe of the same source fragment.</summary>
    private void RecordColumnSourceFragmentSize(IElement element, HtmlRenderFlowBlock completed,
        HtmlRenderBoxStyle style) {
        if (_columnSourceFlowContext == null || style.Display is "inline" or "contents"
            || !TryGetColumnSourceFragment(element, out ColumnSourceFragment fragment)) return;
        CheckCancellation();
        ChargeLayoutOperation("column source box-size progress");
        double boxHeight = Math.Max(0D, completed.Height - style.MarginTop - style.MarginBottom);
        _columnSourceFlowContext.Record(fragment, boxHeight);
    }

    private sealed class ColumnSourceFlowContext {
        private readonly Dictionary<IElement, double> _totals = new();
        internal Dictionary<ColumnSourceFragment, double> Sizes { get; } = new();
        internal HashSet<PositionedElementRequest> PaintedRequests { get; } = new();
        internal HashSet<IElement> FirstFragmentPositionedElements { get; } = new();
        internal Dictionary<IElement, ColumnSourceOpacityOwner> OpacityOwners { get; } = new();

        internal double ConsumedBefore(ColumnSourceFragment fragment) {
            _totals.TryGetValue(fragment.Element, out double total);
            Sizes.TryGetValue(fragment, out double current);
            return Math.Max(0D, total - current);
        }

        internal void Record(ColumnSourceFragment fragment, double height) {
            _totals.TryGetValue(fragment.Element, out double total);
            Sizes.TryGetValue(fragment, out double previous);
            Sizes[fragment] = height;
            _totals[fragment.Element] = Math.Max(0D, total - previous + height);
        }
    }

    private sealed class ColumnSourceFlowScope : IDisposable {
        private readonly HtmlRenderLayoutEngine _owner;
        private readonly ColumnSourceFlowContext? _previous;
        private bool _disposed;

        internal ColumnSourceFlowScope(HtmlRenderLayoutEngine owner) {
            _owner = owner;
            _previous = owner._columnSourceFlowContext;
            owner._columnSourceFlowContext = new ColumnSourceFlowContext();
        }

        public void Dispose() {
            if (_disposed) return;
            _owner._columnSourceFlowContext = _previous;
            _disposed = true;
        }
    }
}
