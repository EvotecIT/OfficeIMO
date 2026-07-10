using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddBorderPaint(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii,
        IElement source,
        string sourceDescription) {
        if (style.UnsupportedBorderPaint.Length > 0) {
            ReportStrokeFallback(
                _reportedBorderPaintFallbacks,
                HtmlRenderDiagnosticCodes.BorderPaintValueUnsupported,
                "A CSS border paint declaration used no-border fallback.",
                style.UnsupportedBorderPaint,
                source,
                sourceDescription);
            return;
        }
        if (style.BorderWidth <= 0D || style.BorderStyle == "none" || style.BorderStyle == "hidden") return;
        if (style.BorderStyle == "double") {
            double strokeWidth = Math.Max(0.01D, style.BorderWidth / 3D);
            AddStrokeVisual(visuals, x, y, width, height, radii, style.BorderColor, strokeWidth, "solid", sourceDescription + ":border-outer");
            double inset = style.BorderWidth * 2D / 3D;
            AddExpandedStrokeVisual(visuals, x, y, width, height, radii, style.BorderColor, strokeWidth, "solid", -inset, sourceDescription + ":border-inner");
            return;
        }
        AddStrokeVisual(visuals, x, y, width, height, radii, style.BorderColor, style.BorderWidth, style.BorderStyle, sourceDescription);
    }

    private void AddOutlinePaint(
        ICollection<HtmlRenderVisual> visuals,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii,
        IElement source,
        string sourceDescription) {
        if (style.UnsupportedOutlinePaint.Length > 0) {
            ReportStrokeFallback(
                _reportedOutlinePaintFallbacks,
                HtmlRenderDiagnosticCodes.OutlinePaintValueUnsupported,
                "A CSS outline paint declaration was omitted.",
                style.UnsupportedOutlinePaint,
                source,
                sourceDescription);
            return;
        }
        if (style.OutlineWidth <= 0D || style.OutlineStyle == "none" || style.OutlineStyle == "hidden") return;
        if (style.OutlineStyle == "double") {
            double strokeWidth = Math.Max(0.01D, style.OutlineWidth / 3D);
            AddExpandedStrokeVisual(visuals, x, y, width, height, radii, style.OutlineColor, strokeWidth, "solid", style.OutlineOffset + style.OutlineWidth / 6D, sourceDescription + ":outline-inner");
            AddExpandedStrokeVisual(visuals, x, y, width, height, radii, style.OutlineColor, strokeWidth, "solid", style.OutlineOffset + style.OutlineWidth * 5D / 6D, sourceDescription + ":outline-outer");
            return;
        }
        AddExpandedStrokeVisual(
            visuals,
            x,
            y,
            width,
            height,
            radii,
            style.OutlineColor,
            style.OutlineWidth,
            style.OutlineStyle,
            style.OutlineOffset + style.OutlineWidth / 2D,
            sourceDescription + ":outline");
    }

    private void ReportStrokeFallback(
        ISet<string> reported,
        string code,
        string message,
        string detail,
        IElement source,
        string sourceDescription) {
        if (!reported.Add(sourceDescription)) return;
        _diagnostics.Add(ComponentName, code, message, HtmlDiagnosticSeverity.Warning, HtmlRenderStyleResolver.DescribeSource(source), detail);
    }

    private static void AddStrokeVisual(
        ICollection<HtmlRenderVisual> visuals,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii,
        OfficeColor color,
        double strokeWidth,
        string style,
        string source) {
        OfficeShape shape = CreateBoxShape(width, height, radii);
        shape.FillColor = null;
        shape.StrokeColor = color;
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = MapStrokeDashStyle(style);
        visuals.Add(new HtmlRenderShape(shape, x, y, visuals.Count, source: source));
    }

    private static void AddExpandedStrokeVisual(
        ICollection<HtmlRenderVisual> visuals,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii,
        OfficeColor color,
        double strokeWidth,
        string style,
        double expansion,
        string source) {
        double targetWidth = width + expansion * 2D;
        double targetHeight = height + expansion * 2D;
        if (targetWidth <= 0.01D || targetHeight <= 0.01D) return;
        OfficeShape shape = CreateBoxShape(targetWidth, targetHeight, radii.Expand(expansion, targetWidth, targetHeight));
        shape.FillColor = null;
        shape.StrokeColor = color;
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = MapStrokeDashStyle(style);
        visuals.Add(new HtmlRenderShape(shape, x - expansion, y - expansion, visuals.Count, source: source));
    }

    private static OfficeStrokeDashStyle MapStrokeDashStyle(string style) => style switch {
        "dashed" => OfficeStrokeDashStyle.Dash,
        "dotted" => OfficeStrokeDashStyle.Dot,
        _ => OfficeStrokeDashStyle.Solid
    };
}
