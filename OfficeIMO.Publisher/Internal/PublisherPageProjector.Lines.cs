using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    /// <summary>Maps native line semantics into the shared scene; marker dimensions remain an explicit approximation.</summary>
    private void ProjectLineDetails(PublisherEscherShape source, OfficeShape shape) {
        if (!shape.StrokeColor.HasValue || shape.StrokeWidth <= 0) return;
        OfficeArtShapeStyle style = source.Style;
        ProjectLineDash(style.LineDashing.GetValueOrDefault(), source.Id, shape);
        shape.StrokeLineCap = style.LineEndCapStyle switch {
            0 => OfficeStrokeLineCap.Round, 1 => OfficeStrokeLineCap.Square, _ => OfficeStrokeLineCap.Butt
        };
        shape.StrokeLineJoin = style.LineJoinStyle switch {
            0 => OfficeStrokeLineJoin.Bevel, 1 => OfficeStrokeLineJoin.Miter, _ => OfficeStrokeLineJoin.Round
        };
        shape.StrokeMiterLimit = style.LineMiterLimit >= 1 ? style.LineMiterLimit.Value : 8;
        if (style.LineEndCapStyle > 2 || style.LineJoinStyle > 2
            || (shape.StrokeLineJoin == OfficeStrokeLineJoin.Miter && style.LineMiterLimit < 1))
            _context.Add("PUB_LINE_STYLE_APPROXIMATED", "An unsupported native cap, join, or miter limit uses its OfficeArt default.",
                OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(source.Id));
        if (style.LineType.GetValueOrDefault() != 0 || style.LineStyle.GetValueOrDefault() != 0)
            _context.Add("PUB_LINE_DETAIL_UNASSESSED", "Native non-solid or compound strokes use a single solid-colored line.",
                OfficeConversionLossKind.Unassessed, PublisherEscherReader.ShapeLocation(source.Id));
        shape.StrokeStartMarker = ProjectLineMarker(style.LineStartArrowhead, style.LineStartArrowWidth, style.LineStartArrowLength, source, shape);
        shape.StrokeEndMarker = ProjectLineMarker(style.LineEndArrowhead, style.LineEndArrowWidth, style.LineEndArrowLength, source, shape);
        if (shape.StrokeStartMarker != null || shape.StrokeEndMarker != null)
            _context.Add("PUB_LINE_MARKER_APPROXIMATED", "Native line marker kinds and relative sizes use shared geometry and stroke-relative dimensions; Publisher-rendered sizing is not qualified.",
                OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(source.Id));
    }

    private OfficeLineMarker? ProjectLineMarker(uint? type, uint? width, uint? length, PublisherEscherShape source, OfficeShape shape) {
        // MSOLINEEND reserves 6 and 7 as decorations that must be ignored.
        OfficeLineMarkerKind kind = type switch {
            null or 0 or 6 or 7 => OfficeLineMarkerKind.None,
            1 => OfficeLineMarkerKind.Triangle, 2 => OfficeLineMarkerKind.Stealth,
            3 => OfficeLineMarkerKind.Diamond, 4 => OfficeLineMarkerKind.Oval, 5 => OfficeLineMarkerKind.Arrow,
            _ => OfficeLineMarkerKind.None
        };
        if (type > 7) _context.Add("PUB_LINE_MARKER_UNSUPPORTED", "An unknown native line decoration was omitted.",
            OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(source.Id));
        if (kind == OfficeLineMarkerKind.None) return null;
        if (shape.Kind != OfficeShapeKind.Line) {
            _context.Add("PUB_LINE_MARKER_GEOMETRY_UNSUPPORTED", "Native line decorations on this geometry remain unassessed and were omitted.",
                OfficeConversionLossKind.Omission, PublisherEscherReader.ShapeLocation(source.Id));
            return null;
        }
        if (width > 2 || length > 2) _context.Add("PUB_LINE_MARKER_SIZE_APPROXIMATED", "An unknown native marker size uses the medium stroke-relative dimensions.",
            OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(source.Id));
        // Use the existing OfficeIMO drawing convention for narrow/medium/wide
        // and short/medium/long ends. Invalid sizes fall back to native medium.
        double widthFactor = width switch { 0 => 3, 2 => 6, _ => 4.5 };
        double lengthFactor = length switch { 0 => 4, 2 => 8, _ => 6 };
        return new OfficeLineMarker(kind, Math.Max(1, shape.StrokeWidth * widthFactor), Math.Max(1, shape.StrokeWidth * lengthFactor));
    }

    private void ProjectLineDash(uint value, uint id, OfficeShape shape) {
        shape.StrokeDashStyle = value switch {
            1 or 6 or 7 => OfficeStrokeDashStyle.Dash, 2 or 5 => OfficeStrokeDashStyle.Dot,
            3 or 8 or 9 => OfficeStrokeDashStyle.DashDot, 4 or 10 => OfficeStrokeDashStyle.DashDotDot,
            _ => OfficeStrokeDashStyle.Solid
        };
        // MSOLINEDASHING's run lengths retain native dot/dash order and spacing.
        // System/device spacing and Publisher's rendered output remain unqualified.
        double[] pattern = value switch {
            1 => new double[] { 3, 1 }, 2 => new double[] { 1, 1 },
            3 => new double[] { 3, 1, 1, 1 }, 4 => new double[] { 3, 1, 1, 1, 1, 1 },
            5 => new double[] { 1, 3 }, 6 => new double[] { 4, 3 }, 7 => new double[] { 8, 3 },
            8 => new double[] { 4, 3, 1, 3 }, 9 => new double[] { 8, 3, 1, 3 },
            10 => new double[] { 8, 3, 1, 3, 1, 3 }, _ => Array.Empty<double>()
        };
        shape.SetStrokeDashArray(pattern.Select(length => length * shape.StrokeWidth));
        if (value != 0) _context.Add("PUB_LINE_DASH_APPROXIMATED", "Native dash order and spacing use stroke-relative lengths; device-dependent spacing and Publisher rendering are not qualified.",
            OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(id));
    }
}
