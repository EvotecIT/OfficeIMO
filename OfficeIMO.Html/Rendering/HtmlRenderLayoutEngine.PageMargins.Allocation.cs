namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Recalculates the complete strip for maximum and then minimum variable-axis content constraints.</summary>
    private (double Left, double Center, double Right) ResolveConstrainedPageMarginStrip(
        HtmlCssPageMarginTemplate? leftBox, HtmlCssPageMarginTemplate? centerBox, HtmlCssPageMarginTemplate? rightBox,
        PageMarginIntrinsicSize leftSize, PageMarginIntrinsicSize centerSize, PageMarginIntrinsicSize rightSize,
        double availableExtent, bool vertical = false) {
        var left = new PageMarginAxisState(leftBox, leftSize, availableExtent, vertical);
        var center = new PageMarginAxisState(centerBox, centerSize, availableExtent, vertical);
        var right = new PageMarginAxisState(rightBox, rightSize, availableExtent, vertical);
        var result = Resolve();
        bool changed = false;
        ApplyConstraint(ref left, result.Left, maximum: true);
        ApplyConstraint(ref center, result.Center, maximum: true);
        ApplyConstraint(ref right, result.Right, maximum: true);
        if (changed) result = Resolve();
        changed = false;
        ApplyConstraint(ref left, result.Left, maximum: false);
        ApplyConstraint(ref center, result.Center, maximum: false);
        ApplyConstraint(ref right, result.Right, maximum: false);
        if (changed) result = Resolve();
        return result;

        (double Left, double Center, double Right) Resolve() {
            CheckCancellation();
            ChargeLayoutOperation(vertical ? "@page side margin-box height allocation" : "@page horizontal margin-box width allocation");
            if (!center.Generated) {
                (double a, double c) = DistributePageMarginAxisPair(left, right, availableExtent);
                return (a, 0D, c);
            }
            var sides = new PageMarginAxisState(left.Generated || right.Generated,
                2D * Math.Max(left.OuterMinimum, right.OuterMinimum),
                2D * Math.Max(left.OuterMaximum, right.OuterMaximum),
                left.OuterExtent.HasValue && right.OuterExtent.HasValue
                    ? 2D * Math.Max(left.OuterExtent.Value, right.OuterExtent.Value) : null,
                2D * Math.Max(left.OuterEdges, right.OuterEdges));
            (double b, _) = DistributePageMarginAxisPair(center, sides, availableExtent);
            double sideExtent = (availableExtent - b) / 2D;
            return (left.Generated ? left.OuterExtent ?? Math.Max(left.OuterEdges, sideExtent) : 0D,
                b, right.Generated ? right.OuterExtent ?? Math.Max(right.OuterEdges, sideExtent) : 0D);
        }

        void ApplyConstraint(ref PageMarginAxisState state, double outerExtent, bool maximum) {
            double? limit = maximum ? state.MaximumConstraint : state.MinimumConstraint;
            if (!state.Generated || !limit.HasValue) return;
            double contentExtent = Math.Max(0D, outerExtent - state.Margins - state.Insets);
            if (maximum ? contentExtent <= limit.Value : contentExtent >= limit.Value) return;
            state = state.WithContentExtent(limit.Value);
            changed = true;
        }
    }

    /// <summary>Distributes signed outer extents while flooring only the enclosed content and border box.</summary>
    private static (double First, double Second) DistributePageMarginAxisPair(
        PageMarginAxisState first, PageMarginAxisState second, double availableExtent) {
        if (first.OuterExtent.HasValue && second.OuterExtent.HasValue)
            return (first.OuterExtent.Value, second.OuterExtent.Value);
        if (first.OuterExtent.HasValue)
            return (first.OuterExtent.Value, Math.Max(second.OuterEdges, availableExtent - first.OuterExtent.Value));
        if (second.OuterExtent.HasValue)
            return (Math.Max(first.OuterEdges, availableExtent - second.OuterExtent.Value), second.OuterExtent.Value);
        (double a, double b) = DistributePageMarginPair(
            (first.Generated, first.OuterMinimum, first.OuterMaximum),
            (second.Generated, second.OuterMinimum, second.OuterMaximum), availableExtent);
        return (Math.Max(first.OuterEdges, a), Math.Max(second.OuterEdges, b));
    }

    private readonly struct PageMarginAxisState {
        internal PageMarginAxisState(HtmlCssPageMarginTemplate? box, PageMarginIntrinsicSize size, double availableExtent, bool vertical) {
            Generated = size.Generated;
            HtmlRenderBoxStyle? style = box?.Typography;
            if (Generated && style != null && !vertical && style.HasIntrinsicWidths)
                style = ApplyIntrinsicWidthValues(style, size.Minimum, size.Maximum, availableExtent);
            Insets = Generated && style != null ? vertical ? style.VerticalInsets : style.HorizontalInsets : 0D;
            // Auto margins have used value zero on the variable axis.
            Margins = Generated && style != null ? vertical
                ? (style.MarginTopAuto ? 0D : style.MarginTop) + (style.MarginBottomAuto ? 0D : style.MarginBottom)
                : (style.MarginLeftAuto ? 0D : style.MarginLeft) + (style.MarginRightAuto ? 0D : style.MarginRight) : 0D;
            OuterEdges = Insets + Margins;
            // Variable-axis allocation uses content coordinates. The shared intrinsic
            // owner has already converted keywords into the selected box edges.
            double borderBoxInsets = style?.BorderBox == true ? Insets : 0D;
            double? ContentExtent(double? value) => value.HasValue
                ? Math.Max(0D, value.Value - borderBoxInsets) : null;
            double? authored = vertical && style != null
                ? style.PageMarginIntrinsicHeight ? size.Maximum + borderBoxInsets : style.ExplicitHeight
                : style?.ExplicitWidth;
            double? width = Generated ? ContentExtent(authored) : 0D;
            OuterExtent = width.HasValue ? width.Value + Insets + Margins : null;
            OuterMinimum = (width ?? size.Minimum) + Insets + Margins;
            OuterMaximum = (width ?? size.Maximum) + Insets + Margins;
            MinimumConstraint = Generated ? ContentExtent(vertical && style != null
                ? style.PageMarginIntrinsicMinHeight ? size.Minimum + borderBoxInsets : style.MinHeight : style?.MinWidth) : null;
            MaximumConstraint = Generated ? ContentExtent(vertical && style != null
                ? style.PageMarginIntrinsicMaxHeight ? size.Maximum + borderBoxInsets : style.MaxHeight : style?.MaxWidth) : null;
        }

        internal PageMarginAxisState(bool generated, double minimum, double maximum, double? width, double outerEdges) {
            Generated = generated;
            Insets = Margins = 0D;
            OuterEdges = outerEdges;
            OuterExtent = generated ? width : 0D;
            OuterMinimum = minimum;
            OuterMaximum = maximum;
            MinimumConstraint = MaximumConstraint = null;
        }

        private PageMarginAxisState(PageMarginAxisState original, double contentExtent) {
            Generated = original.Generated;
            Insets = original.Insets;
            Margins = original.Margins;
            OuterEdges = original.OuterEdges;
            OuterExtent = contentExtent + Insets + Margins;
            OuterMinimum = OuterMaximum = OuterExtent.Value;
            MinimumConstraint = original.MinimumConstraint;
            MaximumConstraint = original.MaximumConstraint;
        }

        internal bool Generated { get; }
        internal double Insets { get; }
        internal double Margins { get; }
        internal double OuterEdges { get; }
        internal double? OuterExtent { get; }
        internal double OuterMinimum { get; }
        internal double OuterMaximum { get; }
        internal double? MinimumConstraint { get; }
        internal double? MaximumConstraint { get; }
        internal PageMarginAxisState WithContentExtent(double width) => new PageMarginAxisState(this, width);
    }
}
