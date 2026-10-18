using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Every set preceding a spanner balances. Only the terminal set follows
    /// column-fill, with a declared height shared by all preceding sets.
    /// </summary>
    private double ResolveSpanningColumnHeight(IReadOnlyList<HtmlRenderFlowBlock> children,
        int count, HtmlRenderBoxStyle style, double precedingHeight, bool finalPartition) {
        double height = finalPartition && style.ColumnFill == "auto"
            ? ResolveNaturalColumnRunHeight(children) : ResolveBalancedColumnHeight(children, count);
        double? declared = ResolveDeclaredColumnContentHeight(style);
        double? maximum = style.MaxHeight.HasValue
            ? Math.Max(style.MaxHeight.Value, style.MinHeight ?? 0D)
                - (style.BorderBox ? style.VerticalInsets : 0D) : null;
        double? remaining = declared.HasValue ? declared.Value - precedingHeight
            : maximum.HasValue ? maximum.Value - precedingHeight : null;
        if (remaining.HasValue) {
            height = finalPartition && style.ColumnFill == "auto" && declared.HasValue
                ? remaining.Value : Math.Min(height, remaining.Value);
        }
        return Math.Max(0.01D, height);
    }

    /// <summary>
    /// Reuses the ordinary physical column bands at the terminal set's actual
    /// origin. The owner retains its original context so rebuilding does not
    /// subtract the preceding sets twice.
    /// </summary>
    private ColumnPageBands? ResolveSpanningColumnPageBands(IElement owner, double width,
        HtmlRenderBoxStyle style, int count, double sourceHeight, double precedingHeight,
        out ColumnPaginationContext? ownerContext) {
        ColumnPageBands? ownerBands = ResolveColumnPageBands(owner, width,
            ResolveDeclaredColumnContentHeight(style), count, sourceHeight);
        ownerContext = ownerBands?.Context;
        if (ownerContext == null) return null;
        double remaining = ownerBands!.FirstHeight - precedingHeight;
        int pageAdvance = 0;
        if (remaining <= 0.0001D) {
            double pages = 1D + Math.Floor(Math.Max(0D, -remaining) / ownerContext.PageHeight);
            if (pages > _options.MaxPageCount - ownerContext.PageNumber)
                throw new InvalidOperationException("HTML rendering exceeded the configured maximum page count.");
            pageAdvance = (int)pages;
            remaining += pageAdvance * ownerContext.PageHeight;
        }
        var context = ownerContext with {
            Boundary = new PagedFloatBoundary(remaining, ownerContext.PageHeight),
            PageNumber = ownerContext.PageNumber + pageAdvance
        };
        return new ColumnPageBands(this, context, count, sourceHeight, HtmlRenderStyleResolver.DescribeSource(owner));
    }
}
