using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private double ResolveFlexBasis(FlexItem item, double availableWidth, int intrinsicDepth = 0, IReadOnlyList<IntrinsicTextRun>? resolvedRuns = null) {
        HtmlRenderBoxStyle style = item.Style;
        double boxBasis;
        if (style.FlexBasis != "auto") {
            if (TryResolveLength(style.FlexBasis, availableWidth, style, out double parsed)) {
                boxBasis = Math.Max(0D, parsed) + (style.BorderBox ? 0D : style.HorizontalInsets);
            } else {
                ReportUnsupportedFlexValue(item, "flex-basis=" + style.FlexBasis);
                boxBasis = ResolveFlexAutoBoxBasis(item, availableWidth, intrinsicDepth, resolvedRuns);
            }
        } else {
            boxBasis = ResolveFlexAutoBoxBasis(item, availableWidth, intrinsicDepth, resolvedRuns);
        }

        return Math.Max(0D, boxBasis + style.MarginLeft + style.MarginRight);
    }

    private double ResolveFlexAutoBoxBasis(FlexItem item, double availableWidth, int intrinsicDepth, IReadOnlyList<IntrinsicTextRun>? resolvedRuns = null) {
        HtmlRenderBoxStyle style = item.Style;
        string tag = item.TagName;
        if (IsReplacedImageElementTag(tag) && item.Element != null) return ResolveReplacedImageBoxWidth(item.Element, style);
        if (IsFormControlElement(tag) && item.Element != null && !UsesButtonChildLayout(item.Element)) {
            double outerWidth = intrinsicDepth > 0
                ? ResolveFormControlIntrinsicOuterWidth(item.Element, style, availableWidth)
                : ResolveFormControlOuterWidth(item.Element, style, availableWidth);
            return Math.Max(0D, outerWidth - style.MarginLeft - style.MarginRight);
        }
        if (style.ExplicitWidth.HasValue) {
            return style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
        }

        if (tag == "table") return availableWidth;
        if (item.Element != null
            && style.Display == "flex"
            && (style.FlexDirection == "row" || style.FlexDirection == "row-reverse")) {
            EnsureDepth(intrinsicDepth + 1, item.Element);
            if (TryCollectFlexItems(item.Element, availableWidth, style, intrinsicDepth + 1,
                    captureRunningElements: false, out List<FlexItem> nestedItems, out _, registerOutOfFlowElements: false)) {
                double nestedWidth = nestedItems.Sum(child => ResolveFlexIntrinsicItemWidth(child, availableWidth, intrinsicDepth + 1))
                    + style.ColumnGap * Math.Max(0, nestedItems.Count - 1);
                return Math.Min(availableWidth, nestedWidth + style.HorizontalInsets);
            }
        }
        IReadOnlyList<IntrinsicTextRun> runs = resolvedRuns ?? ResolveInFlowIntrinsicTextRuns(item, availableWidth, includeDescendantInsets: true);
        double measured = runs.Count == 0 ? 0D : MeasureMaxContentRuns(runs);
        if (item.Element != null) {
            foreach (IElement child in item.Element.Children) {
                if (IsClosedDisclosureChild(child)) continue;
                HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, availableWidth, style);
                if (childStyle.Display == "none" || childStyle.Position == "absolute" || childStyle.Position == "fixed"
                    || !HtmlRenderStyleResolver.IsBlockElement(child, childStyle)) continue;
                if (childStyle.Display == "flex" && childStyle.FlexDirection is "row" or "row-reverse") {
                    HtmlRenderBoxStyle intrinsicStyle = childStyle;
                    if (intrinsicStyle.ExplicitWidthUsesPercentage) {
                        intrinsicStyle = intrinsicStyle.Clone();
                        intrinsicStyle.ExplicitWidth = null;
                        intrinsicStyle.ExplicitWidthUsesPercentage = false;
                    }
                    measured = Math.Max(measured,
                        ResolveFlexAutoBoxBasis(new FlexItem(child, intrinsicStyle, 0), availableWidth, intrinsicDepth + 1)
                        + intrinsicStyle.MarginLeft + intrinsicStyle.MarginRight);
                    continue;
                }
            }
        }
        return measured + style.HorizontalInsets;
    }

    private double ResolveFlexIntrinsicItemWidth(FlexItem item, double availableWidth, int intrinsicDepth) {
        NormalizeFlexIntrinsicConstraints(item);
        double basis = ResolveFlexBasis(item, availableWidth, intrinsicDepth);
        double maxContent = ResolveFlexAutoBoxBasis(item, availableWidth, intrinsicDepth)
            + item.Style.MarginLeft + item.Style.MarginRight;
        item.AutomaticMinimumMainSize = ResolveFlexAutomaticMinimumWidth(item, availableWidth);
        return ClampFlexMainSize(item, Math.Max(basis, maxContent), vertical: false);
    }

    private static void NormalizeFlexIntrinsicConstraints(FlexItem item) {
        HtmlRenderBoxStyle style = item.Style;
        if (style.ExplicitWidthUsesPercentage || style.MaxWidthUsesPercentage
            || style.MinWidthWithIndefiniteReference.HasValue) {
            // Cyclic percentages cannot size their own indefinite flex parent.
            // Keep the absolute contribution of a minimum and definite maxima.
            item.Style = style.Clone();
            if (style.ExplicitWidthUsesPercentage) {
                item.Style.ExplicitWidth = null;
                item.Style.ExplicitWidthUsesPercentage = false;
            }
            if (style.MaxWidthUsesPercentage) {
                item.Style.MaxWidth = null;
                item.Style.MaxWidthUsesPercentage = false;
            }
            if (style.MinWidthWithIndefiniteReference.HasValue) {
                item.Style.MinWidth = style.MinWidthWithIndefiniteReference;
            }
        }
    }

    private double ResolveFlexAutomaticMinimumWidth(FlexItem item, double availableWidth) {
        HtmlRenderBoxStyle style = item.Style;
        if (style.MinWidth.HasValue || style.OverflowX is not ("visible" or "clip")) return 0D;

        double minimum;
        if (IsReplacedImageElementTag(item.TagName) && item.Element != null) {
            minimum = ResolveReplacedImageBoxWidth(item.Element, style);
        } else {
            IReadOnlyList<IntrinsicTextRun> runs = ResolveInFlowIntrinsicTextRuns(item, availableWidth);
            double content = runs.Count == 0 ? 0D : MeasureMinContentRuns(runs);
            content = Math.Max(content, ResolveDescendantReplacedGridContribution(item, availableWidth, minimum: true));
            content = Math.Max(content, ResolveDescendantDefiniteFlexWidth(item, availableWidth));
            minimum = content + style.HorizontalInsets;
        }
        if (style.ExplicitWidth.HasValue)
            minimum = Math.Min(minimum, style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        if (style.MaxWidth.HasValue)
            minimum = Math.Min(minimum, style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
        return Math.Max(0D, minimum + style.MarginLeft + style.MarginRight);
    }

    private double ResolveDescendantDefiniteFlexWidth(FlexItem item, double availableWidth) =>
        item.Element == null ? 0D : ResolveDescendantDefiniteFlexWidth(item.Element, item.Style, availableWidth, 1);

    private double ResolveDescendantDefiniteFlexWidth(IElement parent, HtmlRenderBoxStyle parentStyle, double availableWidth, int depth) {
        double maximum = 0D;
        foreach (IElement child in parent.Children) {
            EnsureDepth(depth, child);
            if (IsClosedDisclosureChild(child) || ShouldSkipElement(child)) continue;
            HtmlRenderBoxStyle childStyle = _styleResolver.Resolve(child, availableWidth, parentStyle);
            if (childStyle.Display == "none" || childStyle.Position == "absolute" || childStyle.Position == "fixed") continue;
            if (IsReplacedImageElement(child)) continue;

            double contribution;
            if (childStyle.ExplicitWidth.HasValue && !childStyle.ExplicitWidthUsesPercentage && childStyle.Display != "inline") {
                contribution = ResolveBoxWidth(availableWidth, childStyle) + childStyle.MarginLeft + childStyle.MarginRight;
            } else {
                double descendant = ResolveDescendantDefiniteFlexWidth(child, childStyle, availableWidth, depth + 1);
                contribution = descendant > 0D ? ResolveGridMeasuredContribution(childStyle, descendant) : 0D;
            }
            maximum = Math.Max(maximum, contribution);
        }
        return maximum;
    }

    private static string CollapseFlexText(string value) {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        return string.Join(" ", value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

}
