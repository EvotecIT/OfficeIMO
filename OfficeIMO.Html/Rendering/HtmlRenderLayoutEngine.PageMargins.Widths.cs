namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Recalculates the complete strip for maximum and then minimum content-width constraints.</summary>
    private (double Left, double Center, double Right) ResolveConstrainedPageMarginStrip(
        HtmlCssPageMarginTemplate? leftBox, HtmlCssPageMarginTemplate? centerBox, HtmlCssPageMarginTemplate? rightBox,
        PageMarginIntrinsicSize leftSize, PageMarginIntrinsicSize centerSize, PageMarginIntrinsicSize rightSize,
        double availableWidth) {
        var left = new PageMarginWidthState(leftBox, leftSize, availableWidth);
        var center = new PageMarginWidthState(centerBox, centerSize, availableWidth);
        var right = new PageMarginWidthState(rightBox, rightSize, availableWidth);
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
            ChargeLayoutOperation("@page horizontal margin-box width allocation");
            if (!center.Generated) {
                (double a, double c) = DistributePageMarginWidthPair(left, right, availableWidth);
                return (a, 0D, c);
            }
            var sides = new PageMarginWidthState(left.Generated || right.Generated,
                2D * Math.Max(left.OuterMinimum, right.OuterMinimum),
                2D * Math.Max(left.OuterMaximum, right.OuterMaximum),
                left.OuterWidth.HasValue && right.OuterWidth.HasValue
                    ? 2D * Math.Max(left.OuterWidth.Value, right.OuterWidth.Value) : null);
            (double b, _) = DistributePageMarginWidthPair(center, sides, availableWidth);
            double sideWidth = (availableWidth - b) / 2D;
            return (left.Generated ? left.OuterWidth ?? Math.Max(left.Insets, sideWidth) : 0D,
                b, right.Generated ? right.OuterWidth ?? Math.Max(right.Insets, sideWidth) : 0D);
        }

        void ApplyConstraint(ref PageMarginWidthState state, double outerWidth, bool maximum) {
            double? limit = maximum ? state.MaximumConstraint : state.MinimumConstraint;
            if (!state.Generated || !limit.HasValue) return;
            double contentWidth = Math.Max(0D, outerWidth - state.Insets);
            if (maximum ? contentWidth <= limit.Value : contentWidth >= limit.Value) return;
            state = state.WithContentWidth(limit.Value);
            changed = true;
        }
    }

    private static (double First, double Second) DistributePageMarginWidthPair(
        PageMarginWidthState first, PageMarginWidthState second, double availableWidth) {
        if (first.OuterWidth.HasValue && second.OuterWidth.HasValue)
            return (first.OuterWidth.Value, second.OuterWidth.Value);
        if (first.OuterWidth.HasValue)
            return (first.OuterWidth.Value, Math.Max(second.Insets, availableWidth - first.OuterWidth.Value));
        if (second.OuterWidth.HasValue)
            return (Math.Max(first.Insets, availableWidth - second.OuterWidth.Value), second.OuterWidth.Value);
        (double a, double b) = DistributePageMarginPair(
            new PageMarginIntrinsicSize(first.Generated, first.OuterMinimum, first.OuterMaximum),
            new PageMarginIntrinsicSize(second.Generated, second.OuterMinimum, second.OuterMaximum), availableWidth);
        return (Math.Max(first.Insets, a), Math.Max(second.Insets, b));
    }

    private readonly struct PageMarginWidthState {
        internal PageMarginWidthState(HtmlCssPageMarginTemplate? box, PageMarginIntrinsicSize size, double availableWidth) {
            Generated = size.Generated;
            HtmlRenderBoxStyle? style = box?.Typography;
            if (Generated && style != null && style.HasIntrinsicWidths)
                style = ApplyIntrinsicWidthValues(style, size.Minimum, size.Maximum, availableWidth);
            Insets = Generated ? style?.HorizontalInsets ?? 0D : 0D;
            double? width = Generated ? style?.ExplicitWidth : 0D;
            OuterWidth = width.HasValue ? width.Value + Insets : null;
            OuterMinimum = (width ?? size.Minimum) + Insets;
            OuterMaximum = (width ?? size.Maximum) + Insets;
            MinimumConstraint = Generated ? style?.MinWidth : null;
            MaximumConstraint = Generated ? style?.MaxWidth : null;
        }

        internal PageMarginWidthState(bool generated, double minimum, double maximum, double? width) {
            Generated = generated;
            Insets = 0D;
            OuterWidth = generated ? width : 0D;
            OuterMinimum = minimum;
            OuterMaximum = maximum;
            MinimumConstraint = MaximumConstraint = null;
        }

        private PageMarginWidthState(PageMarginWidthState original, double contentWidth) {
            Generated = original.Generated;
            Insets = original.Insets;
            OuterWidth = contentWidth + Insets;
            OuterMinimum = OuterMaximum = OuterWidth.Value;
            MinimumConstraint = original.MinimumConstraint;
            MaximumConstraint = original.MaximumConstraint;
        }

        internal bool Generated { get; }
        internal double Insets { get; }
        internal double? OuterWidth { get; }
        internal double OuterMinimum { get; }
        internal double OuterMaximum { get; }
        internal double? MinimumConstraint { get; }
        internal double? MaximumConstraint { get; }
        internal PageMarginWidthState WithContentWidth(double width) => new PageMarginWidthState(this, width);
    }
}
