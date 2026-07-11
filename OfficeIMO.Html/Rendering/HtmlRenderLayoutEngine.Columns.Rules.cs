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
        string source) {
        if (columnCount <= 1 || height <= 0.0001D || style.ColumnRuleWidth <= 0D
            || style.ColumnRuleStyle == "none" || style.ColumnRuleStyle == "hidden") return;
        OfficeColor color = style.ColumnRuleColor ?? style.Color;
        for (int boundary = 1; boundary < columnCount; boundary++) {
            double centerX = contentX + boundary * (columnWidth + gap) - gap / 2D;
            if (style.ColumnRuleStyle == "double") {
                double strokeWidth = Math.Max(0.01D, style.ColumnRuleWidth / 3D);
                double separation = style.ColumnRuleWidth / 3D;
                AddColumnRuleVisual(visuals, centerX - separation, contentY, height, color, strokeWidth, OfficeStrokeDashStyle.Solid, source);
                AddColumnRuleVisual(visuals, centerX + separation, contentY, height, color, strokeWidth, OfficeStrokeDashStyle.Solid, source);
            } else {
                OfficeStrokeDashStyle dashStyle = style.ColumnRuleStyle == "dashed"
                    ? OfficeStrokeDashStyle.Dash
                    : style.ColumnRuleStyle == "dotted" ? OfficeStrokeDashStyle.Dot : OfficeStrokeDashStyle.Solid;
                AddColumnRuleVisual(visuals, centerX, contentY, height, color, style.ColumnRuleWidth, dashStyle, source);
            }
        }
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
