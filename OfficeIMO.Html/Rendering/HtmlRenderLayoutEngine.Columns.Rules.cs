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
        IReadOnlyList<ColumnRuleBand>? bands = null,
        bool extendToPageFragmentEnd = false) {
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
                AddColumnRuleVisual(visuals, centerX - separation, y, extent, color, strokeWidth, OfficeStrokeDashStyle.Solid, source, extendToPageFragmentEnd);
                AddColumnRuleVisual(visuals, centerX + separation, y, extent, color, strokeWidth, OfficeStrokeDashStyle.Solid, source, extendToPageFragmentEnd);
            } else {
                OfficeStrokeDashStyle dashStyle = style.ColumnRuleStyle == "dashed"
                    ? OfficeStrokeDashStyle.Dash
                    : style.ColumnRuleStyle == "dotted" ? OfficeStrokeDashStyle.Dot : OfficeStrokeDashStyle.Solid;
                AddColumnRuleVisual(visuals, centerX, y, extent, color, style.ColumnRuleWidth, dashStyle, source, extendToPageFragmentEnd);
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
        string source,
        bool extendToPageFragmentEnd) {
        OfficeShape shape = OfficeShape.Line(0D, 0D, 0.0001D, height);
        shape.StrokeColor = color;
        shape.StrokeWidth = width;
        shape.StrokeDashStyle = dashStyle;
        visuals.Add(new HtmlRenderShape(shape, x, y, visuals.Count, source: source + "::column-rule",
            canExtendFragmentColumnRule: extendToPageFragmentEnd));
    }

    /// <summary>Extends only owned rule paint; its original source interval still controls fragment membership.</summary>
    private static HtmlRenderShape ExtendFragmentColumnRule(HtmlRenderShape rule, double paintEnd) {
        OfficeShape original = rule.InnerShape;
        OfficeShape line = OfficeShape.Line(0D, 0D, original.Width, paintEnd - rule.Y);
        line.StrokeColor = original.StrokeColor;
        line.StrokeWidth = original.StrokeWidth;
        line.StrokeDashStyle = original.StrokeDashStyle;
        line.StrokeLineCap = original.StrokeLineCap;
        line.StrokeLineJoin = original.StrokeLineJoin;
        line.StrokeMiterLimit = original.StrokeMiterLimit;
        line.SetStrokeDashArray(original.StrokeDashArray, original.StrokeDashOffset);
        return rule.CopyStackingContextTo(new HtmlRenderShape(line, rule.X, rule.Y,
            rule.PaintOrder, source: rule.Source, layoutY: rule.LayoutY, layoutHeight: rule.LayoutHeight));
    }
}
