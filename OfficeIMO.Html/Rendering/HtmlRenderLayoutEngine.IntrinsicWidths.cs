using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Resolves ordinary box constraints once, before margins and numeric layout.</summary>
    private HtmlRenderBoxStyle ResolveOrdinaryIntrinsicWidths(IElement element, HtmlRenderBoxStyle style, double availableWidth, int depth) {
        if (!style.HasIntrinsicWidths || style.IntrinsicWidthsResolved) return style;
        IReadOnlyList<IntrinsicTextRun> runs = ResolveInFlowIntrinsicTextRuns(
            new FlexItem(element, style, 0), availableWidth, depth, includeDescendantInsets: true);
        double minimum = runs.Count == 0 ? 0D : MeasureMinContentRuns(runs);
        double maximum = runs.Count == 0 ? 0D : Math.Max(minimum, MeasureMaxContentRuns(runs));
        return ApplyIntrinsicWidthValues(style, minimum, maximum, availableWidth);
    }

    private static HtmlRenderBoxStyle ApplyIntrinsicWidthValues(
        HtmlRenderBoxStyle style, double minimum, double maximum, double availableWidth, double? intrinsicAvailable = null) {
        if (!style.HasIntrinsicWidths) return style;
        HtmlRenderBoxStyle resolved = style.Clone();
        resolved.IntrinsicWidthsResolved = true;
        double contentAvailable = intrinsicAvailable ?? Math.Max(0D,
            availableWidth - style.MarginLeft - style.MarginRight - style.HorizontalInsets);
        double Resolve(HtmlRenderIntrinsicWidth value) {
            double content = value.Kind switch {
                HtmlRenderIntrinsicWidthKind.MinContent => minimum,
                HtmlRenderIntrinsicWidthKind.MaxContent => maximum,
                _ => Math.Max(minimum, Math.Min(maximum, value.Limit.HasValue
                    ? Math.Max(0D, value.Limit.Value - (style.BorderBox ? style.HorizontalInsets : 0D))
                    : contentAvailable))
            };
            // Keywords measure the content box regardless of box-sizing. Only
            // a functional fit-content argument uses the authored box edges.
            return content + (style.BorderBox ? style.HorizontalInsets : 0D);
        }
        if (style.IntrinsicWidth is { } width) {
            resolved.ExplicitWidth = Resolve(width);
            resolved.ExplicitWidthUsesPercentage = false;
        }
        if (style.IntrinsicMinWidth is { } min) {
            resolved.MinWidth = Resolve(min);
            resolved.MinWidthWithIndefiniteReference = null;
        }
        if (style.IntrinsicMaxWidth is { } max) {
            resolved.MaxWidth = Resolve(max);
            resolved.MaxWidthUsesPercentage = false;
        }
        return resolved;
    }

    /// <summary>Preserves nested authored sizes while the ancestor's width is indefinite.</summary>
    private (double Minimum, double Maximum) ResolveOrdinaryIntrinsicContributions(
        HtmlRenderBoxStyle style, double minimum, double maximum, double availableWidth) {
        double Contribution(double available, double measured) {
            HtmlRenderBoxStyle resolved = ApplyIntrinsicWidthValues(style, minimum, maximum, availableWidth, available);
            return resolved.ExplicitWidth.HasValue && !resolved.ExplicitWidthUsesPercentage
                ? ResolveBoxWidth(availableWidth, resolved) + resolved.MarginLeft + resolved.MarginRight
                : ResolveGridMeasuredContribution(resolved, measured);
        }
        double low = Contribution(minimum, minimum);
        return (low, Math.Max(low, Contribution(maximum, maximum)));
    }
}
