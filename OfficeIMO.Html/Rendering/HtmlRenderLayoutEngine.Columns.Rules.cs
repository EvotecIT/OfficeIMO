using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static void AddColumnRuleVisuals(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double contentX,
        double contentY,
        double columnWidth,
        double gap,
        int columnCount,
        double height,
        string source,
        IReadOnlyList<ColumnRuleBand>? bands = null) {
        if (columnCount <= 1 || height <= 0.0001D || style.ColumnRuleWidth <= 0D
            || style.ColumnRuleStyle == "none" || style.ColumnRuleStyle == "hidden") return;
        OfficeColor color = style.ColumnRuleColor ?? style.Color;
        if (bands != null) {
            foreach (ColumnRuleBand band in bands) {
                foreach (int boundary in band.Boundaries)
                    AddBoundary(boundary, contentY + band.Y, band.Height);
            }
            return;
        }
        for (int boundary = 1; boundary < columnCount; boundary++) {
            AddBoundary(boundary, contentY, height);
        }

        void AddBoundary(int boundary, double y, double extent) {
            double centerX = contentX + boundary * (columnWidth + gap) - gap / 2D;
            if (style.ColumnRuleStyle == "double") {
                double strokeWidth = Math.Max(0.01D, style.ColumnRuleWidth / 3D);
                double separation = style.ColumnRuleWidth / 3D;
                AddColumnRuleVisual(visuals, centerX - separation, y, extent, color, strokeWidth, OfficeStrokeDashStyle.Solid, source);
                AddColumnRuleVisual(visuals, centerX + separation, y, extent, color, strokeWidth, OfficeStrokeDashStyle.Solid, source);
            } else {
                OfficeStrokeDashStyle dashStyle = style.ColumnRuleStyle == "dashed"
                    ? OfficeStrokeDashStyle.Dash
                    : style.ColumnRuleStyle == "dotted" ? OfficeStrokeDashStyle.Dot : OfficeStrokeDashStyle.Solid;
                AddColumnRuleVisual(visuals, centerX, y, extent, color, style.ColumnRuleWidth, dashStyle, source);
            }
        }
    }

    /// <summary>Occupied adjacent columns and rule extent in one physical band.</summary>
    private sealed class ColumnRuleBand {
        internal ColumnRuleBand(double y, double height, IReadOnlyList<int> boundaries) {
            Y = y;
            Height = height;
            Boundaries = boundaries;
        }
        internal double Y { get; }
        internal double Height { get; }
        internal IReadOnlyList<int> Boundaries { get; }
    }

    private static void AddColumnRuleVisual(
        ICollection<HtmlRenderVisual> visuals,
        double x,
        double y,
        double height,
        OfficeColor color,
        double width,
        OfficeStrokeDashStyle dashStyle,
        string source) {
        OfficeShape shape = OfficeShape.Line(0D, 0D, 0.0001D, height);
        shape.StrokeColor = color;
        shape.StrokeWidth = width;
        shape.StrokeDashStyle = dashStyle;
        visuals.Add(new HtmlRenderShape(shape, x, y, visuals.Count, source: source + "::column-rule"));
    }
}
