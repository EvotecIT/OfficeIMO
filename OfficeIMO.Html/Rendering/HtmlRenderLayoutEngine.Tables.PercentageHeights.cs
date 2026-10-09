using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly HashSet<IElement> _reportedTablePercentageHeights = new HashSet<IElement>();

    /// <summary>Measures percentage tracks as auto before the definite table grid distributes its height.</summary>
    private HtmlRenderBoxStyle PrepareTablePercentageHeight(HtmlRenderBoxStyle style, HtmlRenderBoxStyle tableStyle) {
        if (!tableStyle.ExplicitHeight.HasValue || style.TablePercentageHeight.Length == 0) return style;
        HtmlRenderBoxStyle measured = style.Clone();
        double? absoluteTerm = _styleResolver.ResolveTablePercentageHeight(style, 0D);
        measured.ExplicitHeight = absoluteTerm > 0D ? absoluteTerm : null;
        return measured;
    }

    /// <summary>Allocates row shares after content and absolute minima, without expanding a definite table for percentages alone.</summary>
    private bool ApplyTablePercentageHeights(IReadOnlyList<TableRowLayout> rows, HtmlRenderBoxStyle tableStyle, double spacing) {
        double baseTotal = rows.Sum(row => row.Height);
        double target = Math.Max(baseTotal, ResolveTableMinimumHeight(tableStyle) - tableStyle.VerticalInsets - spacing * (rows.Count + 1));
        var references = new double[rows.Count];
        var autoRows = new List<int>();
        bool hasPercentages = false;
        double referenceTotal = 0D;
        double percentageTotal = 0D;
        for (int index = 0; index < rows.Count; index++) {
            TableRowLayout row = rows[index];
            double requested = 0D;
            bool specified = row.Style.ExplicitHeight.HasValue;
            Include(row.Element, row.Style);
            foreach (TableCellLayout cell in row.Cells) {
                if (cell.RowSpan != 1) continue;
                specified |= cell.Style.ExplicitHeight.HasValue;
                Include(cell.Element, cell.Style);
            }
            references[index] = Math.Max(row.Height, requested);
            referenceTotal += references[index];
            percentageTotal += requested;
            if (!specified) autoRows.Add(index);

            void Include(IElement element, HtmlRenderBoxStyle style) {
                if (style.TablePercentageHeight.Length == 0) return;
                double? height = _styleResolver.ResolveTablePercentageHeight(style, target);
                if (!height.HasValue) {
                    ReportTablePercentageHeightFallback(element, style, "unresolved percentage height calculation");
                    return;
                }
                specified = true;
                hasPercentages = true;
                // A percentage is a share of the usable row grid; cell insets
                // remain inside that share and still constrain the content minimum.
                requested = Math.Max(requested, height.Value);
            }
        }
        if (!hasPercentages) return false;
        if (percentageTotal > target + 0.0001D) {
            foreach (TableRowLayout row in rows) {
                if (row.Style.TablePercentageHeight.Length > 0) ReportTablePercentageHeightFallback(row.Element, row.Style, "oversubscribed percentage row shares");
                foreach (TableCellLayout cell in row.Cells) {
                    if (cell.RowSpan == 1 && cell.Style.TablePercentageHeight.Length > 0) ReportTablePercentageHeightFallback(cell.Element, cell.Style, "oversubscribed percentage row shares");
                }
            }
        }
        if (referenceTotal > target) {
            double fraction = Math.Max(0D, target - baseTotal) / (referenceTotal - baseTotal);
            for (int index = 0; index < rows.Count; index++) {
                rows[index].Height += (references[index] - rows[index].Height) * fraction;
            }
        } else {
            double extra = target - referenceTotal;
            for (int index = 0; index < rows.Count; index++) rows[index].Height = references[index];
            if (autoRows.Count > 0) {
                foreach (int index in autoRows) rows[index].Height += extra / autoRows.Count;
            } else if (referenceTotal > 0D) {
                foreach (TableRowLayout row in rows) row.Height += extra * row.Height / referenceTotal;
            }
        }
        return true;
    }

    private void ReportTablePercentageHeightFallback(IElement element, HtmlRenderBoxStyle style, string reason) {
        if (!_reportedTablePercentageHeights.Add(element)) return;
        _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableValueUnsupported,
            "A table height used an allocation outside the qualified percentage row and cell subset.",
            HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(element),
            "height=" + style.TablePercentageHeight + ";" + reason, OfficeConversionLossKind.Approximation);
    }
}
