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
                sourceDescription,
                HtmlConversionLossKind.Omission);
            return;
        }
        if (!style.Borders.HasPaint) return;
        if (style.Borders.IsUniform) {
            HtmlRenderBorderSide border = style.Borders.Top;
            if (border.Style == "double") {
                double strokeWidth = Math.Max(0.01D, border.Width / 3D);
                AddStrokeVisual(visuals, x, y, width, height, radii, border.Color, strokeWidth, "solid", sourceDescription + ":border-outer");
                double inset = border.Width * 2D / 3D;
                AddExpandedStrokeVisual(visuals, x, y, width, height, radii, border.Color, strokeWidth, "solid", -inset, sourceDescription + ":border-inner");
                return;
            }
            AddStrokeVisual(visuals, x, y, width, height, radii, border.Color, border.Width, border.Style, sourceDescription);
            return;
        }

        AddBorderSidePaint(visuals, HtmlBorderEdge.Top, style.Borders.Top, x, y, width, height, radii, sourceDescription);
        AddBorderSidePaint(visuals, HtmlBorderEdge.Right, style.Borders.Right, x, y, width, height, radii, sourceDescription);
        AddBorderSidePaint(visuals, HtmlBorderEdge.Bottom, style.Borders.Bottom, x, y, width, height, radii, sourceDescription);
        AddBorderSidePaint(visuals, HtmlBorderEdge.Left, style.Borders.Left, x, y, width, height, radii, sourceDescription);
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
                sourceDescription,
                HtmlConversionLossKind.Omission);
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
        string sourceDescription,
        HtmlConversionLossKind lossKind) {
        if (!reported.Add(sourceDescription)) return;
        _diagnostics.Add(
            ComponentName,
            code,
            message,
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(source),
            detail,
            lossKind);
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

    private static void AddBorderSidePaint(
        ICollection<HtmlRenderVisual> visuals,
        HtmlBorderEdge edge,
        HtmlRenderBorderSide border,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii,
        string source) {
        if (!border.IsPainted) return;
        string edgeSource = source + ":border-" + edge.ToString().ToLowerInvariant();
        if (border.Style != "double") {
            AddBorderSideStroke(visuals, edge, border.Color, border.Width, border.Style, x, y, width, height, radii, edgeSource);
            return;
        }

        double strokeWidth = Math.Max(0.01D, border.Width / 3D);
        AddBorderSideStroke(visuals, edge, border.Color, strokeWidth, "solid", x, y, width, height, radii, edgeSource + "-outer");
        double inset = border.Width * 2D / 3D;
        double innerWidth = width - inset * 2D;
        double innerHeight = height - inset * 2D;
        if (innerWidth <= 0.01D || innerHeight <= 0.01D) return;
        HtmlResolvedBorderRadii innerRadii = radii.Inset(inset, inset, inset, inset, innerWidth, innerHeight);
        AddBorderSideStroke(visuals, edge, border.Color, strokeWidth, "solid", x + inset, y + inset, innerWidth, innerHeight, innerRadii, edgeSource + "-inner");
    }

    private static void AddBorderSideStroke(
        ICollection<HtmlRenderVisual> visuals,
        HtmlBorderEdge edge,
        OfficeColor color,
        double strokeWidth,
        string style,
        double x,
        double y,
        double width,
        double height,
        HtmlResolvedBorderRadii radii,
        string source) {
        OfficeShape shape = OfficeShape.Path(CreateBorderSidePath(edge, width, height, radii.Normalize(width, height)));
        shape.FillColor = null;
        shape.StrokeColor = color;
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = MapStrokeDashStyle(style);
        visuals.Add(new HtmlRenderShape(shape, x, y, visuals.Count, source: source));
    }

    private static IReadOnlyList<OfficePathCommand> CreateBorderSidePath(
        HtmlBorderEdge edge,
        double width,
        double height,
        HtmlResolvedBorderRadii radii) {
        const double kappa = 0.5522847498307936D;
        SplitCubic(new CubicSegment(
            new OfficePoint(0D, radii.TopLeftY),
            new OfficePoint(0D, radii.TopLeftY * (1D - kappa)),
            new OfficePoint(radii.TopLeftX * (1D - kappa), 0D),
            new OfficePoint(radii.TopLeftX, 0D)), out CubicSegment topLeftFirst, out CubicSegment topLeftSecond);
        SplitCubic(new CubicSegment(
            new OfficePoint(width - radii.TopRightX, 0D),
            new OfficePoint(width - radii.TopRightX * (1D - kappa), 0D),
            new OfficePoint(width, radii.TopRightY * (1D - kappa)),
            new OfficePoint(width, radii.TopRightY)), out CubicSegment topRightFirst, out CubicSegment topRightSecond);
        SplitCubic(new CubicSegment(
            new OfficePoint(width, height - radii.BottomRightY),
            new OfficePoint(width, height - radii.BottomRightY * (1D - kappa)),
            new OfficePoint(width - radii.BottomRightX * (1D - kappa), height),
            new OfficePoint(width - radii.BottomRightX, height)), out CubicSegment bottomRightFirst, out CubicSegment bottomRightSecond);
        SplitCubic(new CubicSegment(
            new OfficePoint(radii.BottomLeftX, height),
            new OfficePoint(radii.BottomLeftX * (1D - kappa), height),
            new OfficePoint(0D, height - radii.BottomLeftY * (1D - kappa)),
            new OfficePoint(0D, height - radii.BottomLeftY)), out CubicSegment bottomLeftFirst, out CubicSegment bottomLeftSecond);

        var commands = new List<OfficePathCommand> {
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.MoveTo(width, height)
        };
        switch (edge) {
            case HtmlBorderEdge.Top:
                AppendSide(commands, topLeftSecond, topRightFirst);
                break;
            case HtmlBorderEdge.Right:
                AppendSide(commands, topRightSecond, bottomRightFirst);
                break;
            case HtmlBorderEdge.Bottom:
                AppendSide(commands, bottomRightSecond, bottomLeftFirst);
                break;
            default:
                AppendSide(commands, bottomLeftSecond, topLeftFirst);
                break;
        }
        return commands;
    }

    private static void AppendSide(ICollection<OfficePathCommand> commands, CubicSegment first, CubicSegment second) {
        commands.Add(OfficePathCommand.MoveTo(first.Start.X, first.Start.Y));
        commands.Add(OfficePathCommand.CubicBezierTo(first.Control1.X, first.Control1.Y, first.Control2.X, first.Control2.Y, first.End.X, first.End.Y));
        commands.Add(OfficePathCommand.LineTo(second.Start.X, second.Start.Y));
        commands.Add(OfficePathCommand.CubicBezierTo(second.Control1.X, second.Control1.Y, second.Control2.X, second.Control2.Y, second.End.X, second.End.Y));
    }

    private static void SplitCubic(CubicSegment source, out CubicSegment first, out CubicSegment second) {
        OfficePoint p01 = Midpoint(source.Start, source.Control1);
        OfficePoint p12 = Midpoint(source.Control1, source.Control2);
        OfficePoint p23 = Midpoint(source.Control2, source.End);
        OfficePoint p012 = Midpoint(p01, p12);
        OfficePoint p123 = Midpoint(p12, p23);
        OfficePoint midpoint = Midpoint(p012, p123);
        first = new CubicSegment(source.Start, p01, p012, midpoint);
        second = new CubicSegment(midpoint, p123, p23, source.End);
    }

    private static OfficePoint Midpoint(OfficePoint left, OfficePoint right) =>
        new OfficePoint((left.X + right.X) / 2D, (left.Y + right.Y) / 2D);

    private static OfficeStrokeDashStyle MapStrokeDashStyle(string style) => style switch {
        "dashed" => OfficeStrokeDashStyle.Dash,
        "dotted" => OfficeStrokeDashStyle.Dot,
        _ => OfficeStrokeDashStyle.Solid
    };

    private enum HtmlBorderEdge {
        Top,
        Right,
        Bottom,
        Left
    }

    private readonly struct CubicSegment {
        internal CubicSegment(OfficePoint start, OfficePoint control1, OfficePoint control2, OfficePoint end) {
            Start = start;
            Control1 = control1;
            Control2 = control2;
            End = end;
        }

        internal OfficePoint Start { get; }
        internal OfficePoint Control1 { get; }
        internal OfficePoint Control2 { get; }
        internal OfficePoint End { get; }
    }
}
