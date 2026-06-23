using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared SVG formatting helpers used by OfficeIMO renderers.
/// </summary>
public static partial class OfficeSvgFormatting {
    /// <summary>
    /// Formats a numeric SVG attribute value using invariant culture and compact precision.
    /// </summary>
    /// <param name="value">Numeric value to format.</param>
    /// <returns>Formatted SVG number.</returns>
    public static string FormatNumber(double value) {
        if (System.Math.Abs(value) < 0.0000001D) {
            value = 0D;
        }

        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Escapes text for XML/SVG attribute or element content.
    /// </summary>
    /// <param name="value">Text to escape.</param>
    /// <returns>Escaped text, or an empty string for null.</returns>
    public static string Escape(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        return value!
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
    }

    /// <summary>
    /// Converts an Office color to a CSS RGB hex color for SVG paint attributes.
    /// </summary>
    /// <param name="color">Office color.</param>
    /// <returns>CSS color in <c>#RRGGBB</c> form.</returns>
    public static string ToCssColor(OfficeColor color) => "#" + color.ToRgbHex();

    /// <summary>
    /// Converts an Office color alpha channel to an SVG opacity value.
    /// </summary>
    /// <param name="color">Office color.</param>
    /// <returns>Opacity between 0 and 1.</returns>
    public static double ToOpacity(OfficeColor color) => color.A / 255D;

    /// <summary>
    /// Appends an escaped SVG attribute to a markup builder.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="attributeName">Attribute name.</param>
    /// <param name="value">Attribute value.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendAttribute(this StringBuilder builder, string attributeName, string? value) {
        builder.Append(' ')
            .Append(attributeName)
            .Append("=\"")
            .Append(Escape(value))
            .Append('"');
        return builder;
    }

    /// <summary>
    /// Appends a numeric SVG attribute using shared invariant formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="attributeName">Attribute name.</param>
    /// <param name="value">Numeric attribute value.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendNumberAttribute(this StringBuilder builder, string attributeName, double value) =>
        builder.AppendAttribute(attributeName, FormatNumber(value));

    /// <summary>
    /// Extracts the content inside an SVG root element.
    /// </summary>
    /// <param name="svg">SVG markup, or already-inner SVG content.</param>
    /// <returns>Inner SVG markup when a root SVG element is found; otherwise the original value.</returns>
    public static string ExtractSvgInner(string svg) {
        if (string.IsNullOrEmpty(svg)) {
            return string.Empty;
        }

        int start = svg.IndexOf('>');
        int end = svg.LastIndexOf("</svg>", StringComparison.OrdinalIgnoreCase);
        return start >= 0 && end > start ? svg.Substring(start + 1, end - start - 1) : svg;
    }

    /// <summary>
    /// Appends the start tag for a nested SVG viewport.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="x">Nested viewport x-coordinate.</param>
    /// <param name="y">Nested viewport y-coordinate.</param>
    /// <param name="width">Nested viewport width.</param>
    /// <param name="height">Nested viewport height.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendNestedSvgStart(this StringBuilder builder, double x, double y, double width, double height) {
        builder.Append("<svg")
            .AppendNumberAttribute("x", x)
            .AppendNumberAttribute("y", y)
            .AppendNumberAttribute("width", width)
            .AppendNumberAttribute("height", height)
            .AppendAttribute("viewBox", "0 0 " + FormatNumber(width) + " " + FormatNumber(height))
            .Append(">");
        return builder;
    }

    /// <summary>
    /// Appends the end tag for a nested SVG viewport.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendNestedSvgEnd(this StringBuilder builder) {
        builder.Append("</svg>");
        return builder;
    }

    /// <summary>
    /// Appends a nested SVG viewport with the supplied inner content.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="x">Nested viewport x-coordinate.</param>
    /// <param name="y">Nested viewport y-coordinate.</param>
    /// <param name="width">Nested viewport width.</param>
    /// <param name="height">Nested viewport height.</param>
    /// <param name="innerContent">SVG content to place inside the nested viewport.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendNestedSvg(this StringBuilder builder, double x, double y, double width, double height, string innerContent) {
        builder.AppendNestedSvgStart(x, y, width, height);
        builder.Append(innerContent);
        return builder.AppendNestedSvgEnd();
    }

    /// <summary>
    /// Appends an SVG paint attribute and matching opacity attribute when the color is transparent.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="attributeName">Paint attribute name, such as <c>fill</c> or <c>stroke</c>.</param>
    /// <param name="color">Office color.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendPaintAttribute(this StringBuilder builder, string attributeName, OfficeColor color) {
        builder.AppendAttribute(attributeName, ToCssColor(color));
        if (color.A < 255) {
            builder.AppendNumberAttribute(attributeName + "-opacity", ToOpacity(color));
        }

        return builder;
    }

    /// <summary>
    /// Appends a reusable SVG linear-gradient definition.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="id">Gradient identifier.</param>
    /// <param name="gradient">Gradient definition.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendLinearGradientDefinition(this StringBuilder builder, string id, OfficeLinearGradient gradient) {
        if (gradient == null) {
            throw new ArgumentNullException(nameof(gradient));
        }

        builder.Append("<defs><linearGradient id=\"")
            .Append(Escape(id))
            .Append("\" x1=\"")
            .Append(FormatNumber(gradient.StartX * 100D))
            .Append("%\" y1=\"")
            .Append(FormatNumber(gradient.StartY * 100D))
            .Append("%\" x2=\"")
            .Append(FormatNumber(gradient.EndX * 100D))
            .Append("%\" y2=\"")
            .Append(FormatNumber(gradient.EndY * 100D))
            .Append("%\">");

        for (int i = 0; i < gradient.Stops.Count; i++) {
            OfficeGradientStop stop = gradient.Stops[i];
            builder.Append("<stop offset=\"")
                .Append(FormatNumber(stop.Offset * 100D))
                .Append("%\" stop-color=\"")
                .Append(ToCssColor(stop.Color))
                .Append('"');

            double opacity = ToOpacity(stop.Color);
            if (opacity < 1D) {
                builder.AppendNumberAttribute("stop-opacity", opacity);
            }

            builder.Append("/>");
        }

        builder.Append("</linearGradient></defs>");
        return builder;
    }

    /// <summary>
    /// Converts a stroke line cap value to an SVG attribute value.
    /// </summary>
    /// <param name="cap">Stroke line cap.</param>
    /// <returns>SVG stroke-linecap value.</returns>
    public static string FormatStrokeLineCap(OfficeStrokeLineCap cap) {
        switch (cap) {
            case OfficeStrokeLineCap.Round:
                return "round";
            case OfficeStrokeLineCap.Square:
                return "square";
            default:
                return "butt";
        }
    }

    /// <summary>
    /// Converts a stroke line join value to an SVG attribute value.
    /// </summary>
    /// <param name="join">Stroke line join.</param>
    /// <returns>SVG stroke-linejoin value.</returns>
    public static string FormatStrokeLineJoin(OfficeStrokeLineJoin join) {
        switch (join) {
            case OfficeStrokeLineJoin.Bevel:
                return "bevel";
            case OfficeStrokeLineJoin.Round:
                return "round";
            default:
                return "miter";
        }
    }

    /// <summary>
    /// Appends an SVG stroke-linecap attribute.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="cap">Stroke line cap.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendStrokeLineCapAttribute(this StringBuilder builder, OfficeStrokeLineCap cap) =>
        builder.AppendAttribute("stroke-linecap", FormatStrokeLineCap(cap));

    /// <summary>
    /// Appends an SVG stroke-linejoin attribute.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="join">Stroke line join.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendStrokeLineJoinAttribute(this StringBuilder builder, OfficeStrokeLineJoin join) =>
        builder.AppendAttribute("stroke-linejoin", FormatStrokeLineJoin(join));

    /// <summary>
    /// Appends an SVG stroke-dasharray attribute when a dash pattern is present.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="dashArray">SVG dash-array value, or <c>null</c> for solid strokes.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendStrokeDashArrayAttribute(this StringBuilder builder, string? dashArray) {
        if (!string.IsNullOrEmpty(dashArray)) {
            builder.AppendAttribute("stroke-dasharray", dashArray);
        }

        return builder;
    }

    /// <summary>
    /// Appends an SVG stroke-dasharray attribute for a shared Office stroke dash style.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="dashStyle">Office stroke dash style.</param>
    /// <param name="strokeWidth">Rendered stroke width.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendStrokeDashStyleAttribute(this StringBuilder builder, OfficeStrokeDashStyle dashStyle, double strokeWidth) =>
        builder.AppendStrokeDashArrayAttribute(dashStyle.GetSvgDashArray(strokeWidth));

    /// <summary>
    /// Appends an SVG clip-path reference attribute for the supplied clip path identifier.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="clipPathId">Clip path identifier without the leading <c>#</c>.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendClipPathReference(this StringBuilder builder, string clipPathId) =>
        builder.AppendAttribute("clip-path", "url(#" + clipPathId + ")");

    /// <summary>
    /// Formats an SVG rotate transform using shared invariant numeric formatting.
    /// </summary>
    /// <param name="degrees">Rotation angle in degrees.</param>
    /// <param name="centerX">Rotation center x-coordinate.</param>
    /// <param name="centerY">Rotation center y-coordinate.</param>
    /// <returns>SVG rotate transform value.</returns>
    public static string FormatRotateTransform(double degrees, double centerX, double centerY) =>
        "rotate(" + FormatNumber(degrees) + " " + FormatNumber(centerX) + " " + FormatNumber(centerY) + ")";

    /// <summary>
    /// Formats an SVG rotate transform around the current transform origin using shared invariant numeric formatting.
    /// </summary>
    /// <param name="degrees">Rotation angle in degrees.</param>
    /// <returns>SVG rotate transform value.</returns>
    public static string FormatRotateTransform(double degrees) =>
        "rotate(" + FormatNumber(degrees) + ")";

    /// <summary>
    /// Appends an SVG rotate transform attribute using shared invariant numeric formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="degrees">Rotation angle in degrees.</param>
    /// <param name="centerX">Rotation center x-coordinate.</param>
    /// <param name="centerY">Rotation center y-coordinate.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendRotateTransformAttribute(this StringBuilder builder, double degrees, double centerX, double centerY) =>
        builder.AppendAttribute("transform", FormatRotateTransform(degrees, centerX, centerY));

    /// <summary>
    /// Formats an SVG matrix transform using shared invariant numeric formatting.
    /// </summary>
    /// <param name="transform">Affine transform matrix.</param>
    /// <param name="placementX">Additional x placement offset applied to the transform.</param>
    /// <param name="placementY">Additional y placement offset applied to the transform.</param>
    /// <returns>SVG matrix transform value.</returns>
    public static string FormatMatrixTransform(OfficeTransform transform, double placementX = 0D, double placementY = 0D) =>
        "matrix(" +
        FormatNumber(transform.M11) + " " +
        FormatNumber(transform.M12) + " " +
        FormatNumber(transform.M21) + " " +
        FormatNumber(transform.M22) + " " +
        FormatNumber(transform.OffsetX + placementX) + " " +
        FormatNumber(transform.OffsetY + placementY) + ")";

    /// <summary>
    /// Appends an SVG matrix transform attribute using shared invariant numeric formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="transform">Affine transform matrix.</param>
    /// <param name="placementX">Additional x placement offset applied to the transform.</param>
    /// <param name="placementY">Additional y placement offset applied to the transform.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendMatrixTransformAttribute(this StringBuilder builder, OfficeTransform transform, double placementX = 0D, double placementY = 0D) =>
        builder.AppendAttribute("transform", FormatMatrixTransform(transform, placementX, placementY));

    /// <summary>
    /// Appends an SVG rectangular clip path definition.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="clipPathId">Clip path identifier.</param>
    /// <param name="x">Clip rectangle left coordinate.</param>
    /// <param name="y">Clip rectangle top coordinate.</param>
    /// <param name="width">Clip rectangle width.</param>
    /// <param name="height">Clip rectangle height.</param>
    /// <param name="wrapInDefs">When true, wraps the clip path in a <c>defs</c> element.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendRectClipPathDefinition(this StringBuilder builder, string clipPathId, double x, double y, double width, double height, bool wrapInDefs = false) {
        if (wrapInDefs) {
            builder.Append("<defs>");
        }

        builder.Append("<clipPath")
            .AppendAttribute("id", clipPathId)
            .Append("><rect")
            .AppendNumberAttribute("x", x)
            .AppendNumberAttribute("y", y)
            .AppendNumberAttribute("width", width)
            .AppendNumberAttribute("height", height)
            .Append("/></clipPath>");

        if (wrapInDefs) {
            builder.Append("</defs>");
        }

        return builder;
    }

    /// <summary>
    /// Appends an SVG <c>points</c> attribute for point-list based elements.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="points">Point list.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendPointsAttribute(this StringBuilder builder, IReadOnlyList<OfficePoint> points) {
        builder.Append(" points=\"");
        for (int i = 0; i < points.Count; i++) {
            if (i > 0) {
                builder.Append(' ');
            }

            builder.Append(FormatNumber(points[i].X)).Append(',').Append(FormatNumber(points[i].Y));
        }

        builder.Append('"');
        return builder;
    }

    /// <summary>
    /// Appends a complete SVG polyline element using shared point-list formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="points">Polyline points.</param>
    /// <param name="attributes">Optional already-formatted SVG attributes to append before <c>points</c>.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendPolylineElement(this StringBuilder builder, IReadOnlyList<OfficePoint> points, string? attributes = null) {
        builder.Append("<polyline");
        if (!string.IsNullOrEmpty(attributes)) {
            builder.Append(attributes);
        }

        builder.AppendPointsAttribute(points);
        builder.Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete SVG circle element using shared numeric formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="cx">Circle center x-coordinate.</param>
    /// <param name="cy">Circle center y-coordinate.</param>
    /// <param name="radius">Circle radius.</param>
    /// <param name="attributes">Optional already-formatted SVG attributes to append after the geometry.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendCircleElement(this StringBuilder builder, double cx, double cy, double radius, string? attributes = null) {
        builder.Append("<circle")
            .AppendNumberAttribute("cx", cx)
            .AppendNumberAttribute("cy", cy)
            .AppendNumberAttribute("r", radius);
        if (!string.IsNullOrEmpty(attributes)) {
            builder.Append(attributes);
        }

        builder.Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete filled SVG circle element using shared paint formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="cx">Circle center x-coordinate.</param>
    /// <param name="cy">Circle center y-coordinate.</param>
    /// <param name="radius">Circle radius.</param>
    /// <param name="fill">Circle fill color.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendCircleElement(this StringBuilder builder, double cx, double cy, double radius, OfficeColor fill) {
        builder.Append("<circle")
            .AppendNumberAttribute("cx", cx)
            .AppendNumberAttribute("cy", cy)
            .AppendNumberAttribute("r", radius)
            .AppendPaintAttribute("fill", fill)
            .Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete SVG ellipse element using shared numeric formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="cx">Ellipse center x-coordinate.</param>
    /// <param name="cy">Ellipse center y-coordinate.</param>
    /// <param name="rx">Horizontal radius.</param>
    /// <param name="ry">Vertical radius.</param>
    /// <param name="attributes">Optional already-formatted SVG attributes to append after the geometry.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendEllipseElement(this StringBuilder builder, double cx, double cy, double rx, double ry, string? attributes = null) {
        builder.Append("<ellipse")
            .AppendNumberAttribute("cx", cx)
            .AppendNumberAttribute("cy", cy)
            .AppendNumberAttribute("rx", rx)
            .AppendNumberAttribute("ry", ry);
        if (!string.IsNullOrEmpty(attributes)) {
            builder.Append(attributes);
        }

        builder.Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete filled SVG ellipse element using shared paint formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="cx">Ellipse center x-coordinate.</param>
    /// <param name="cy">Ellipse center y-coordinate.</param>
    /// <param name="rx">Horizontal radius.</param>
    /// <param name="ry">Vertical radius.</param>
    /// <param name="fill">Ellipse fill color.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendEllipseElement(this StringBuilder builder, double cx, double cy, double rx, double ry, OfficeColor fill) {
        builder.Append("<ellipse")
            .AppendNumberAttribute("cx", cx)
            .AppendNumberAttribute("cy", cy)
            .AppendNumberAttribute("rx", rx)
            .AppendNumberAttribute("ry", ry)
            .AppendPaintAttribute("fill", fill)
            .Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete SVG rectangle element using shared numeric formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="x">Rectangle x-coordinate.</param>
    /// <param name="y">Rectangle y-coordinate.</param>
    /// <param name="width">Rectangle width.</param>
    /// <param name="height">Rectangle height.</param>
    /// <param name="attributes">Optional already-formatted SVG attributes to append after the geometry.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendRectElement(this StringBuilder builder, double x, double y, double width, double height, string? attributes = null) {
        builder.Append("<rect")
            .AppendNumberAttribute("x", x)
            .AppendNumberAttribute("y", y)
            .AppendNumberAttribute("width", width)
            .AppendNumberAttribute("height", height);
        if (!string.IsNullOrEmpty(attributes)) {
            builder.Append(attributes);
        }

        builder.Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete SVG rounded rectangle element using shared numeric formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="x">Rectangle x-coordinate.</param>
    /// <param name="y">Rectangle y-coordinate.</param>
    /// <param name="width">Rectangle width.</param>
    /// <param name="height">Rectangle height.</param>
    /// <param name="rx">Horizontal corner radius.</param>
    /// <param name="ry">Vertical corner radius.</param>
    /// <param name="attributes">Optional already-formatted SVG attributes to append after the geometry.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendRectElement(this StringBuilder builder, double x, double y, double width, double height, double rx, double ry, string? attributes = null) {
        builder.Append("<rect")
            .AppendNumberAttribute("x", x)
            .AppendNumberAttribute("y", y)
            .AppendNumberAttribute("width", width)
            .AppendNumberAttribute("height", height)
            .AppendNumberAttribute("rx", rx)
            .AppendNumberAttribute("ry", ry);
        if (!string.IsNullOrEmpty(attributes)) {
            builder.Append(attributes);
        }

        builder.Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete SVG line element using shared numeric formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="x1">Line start x-coordinate.</param>
    /// <param name="y1">Line start y-coordinate.</param>
    /// <param name="x2">Line end x-coordinate.</param>
    /// <param name="y2">Line end y-coordinate.</param>
    /// <param name="attributes">Optional already-formatted SVG attributes to append after the coordinates.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendLineElement(this StringBuilder builder, double x1, double y1, double x2, double y2, string? attributes = null) {
        builder.Append("<line")
            .AppendNumberAttribute("x1", x1)
            .AppendNumberAttribute("y1", y1)
            .AppendNumberAttribute("x2", x2)
            .AppendNumberAttribute("y2", y2);
        if (!string.IsNullOrEmpty(attributes)) {
            builder.Append(attributes);
        }

        builder.Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete stroked SVG line element using shared paint, dash, and cap formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="x1">Line start x-coordinate.</param>
    /// <param name="y1">Line start y-coordinate.</param>
    /// <param name="x2">Line end x-coordinate.</param>
    /// <param name="y2">Line end y-coordinate.</param>
    /// <param name="stroke">Stroke color.</param>
    /// <param name="strokeWidth">Stroke width.</param>
    /// <param name="dashStyle">Stroke dash style.</param>
    /// <param name="lineCap">Optional stroke line cap.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendLineElement(this StringBuilder builder, double x1, double y1, double x2, double y2, OfficeColor stroke, double strokeWidth, OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? lineCap = null) {
        builder.Append("<line")
            .AppendNumberAttribute("x1", x1)
            .AppendNumberAttribute("y1", y1)
            .AppendNumberAttribute("x2", x2)
            .AppendNumberAttribute("y2", y2)
            .AppendPaintAttribute("stroke", stroke)
            .AppendNumberAttribute("stroke-width", strokeWidth);
        string? dashArray = dashStyle.GetSvgDashArray(strokeWidth);
        builder.AppendStrokeDashArrayAttribute(dashArray);
        if (lineCap.HasValue) {
            builder.AppendStrokeLineCapAttribute(lineCap.Value);
        }

        builder.Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete SVG polygon element using shared point-list formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="points">Polygon points.</param>
    /// <param name="attributes">Optional already-formatted SVG attributes to append after <c>points</c>.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendPolygonElement(this StringBuilder builder, IReadOnlyList<OfficePoint> points, string? attributes = null) {
        builder.Append("<polygon")
            .AppendPointsAttribute(points);
        if (!string.IsNullOrEmpty(attributes)) {
            builder.Append(attributes);
        }

        builder.Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete SVG polygon element with shared fill/stroke formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="points">Polygon points.</param>
    /// <param name="fill">Polygon fill color.</param>
    /// <param name="stroke">Optional polygon stroke color.</param>
    /// <param name="strokeWidth">Optional polygon stroke width.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendPolygonElement(this StringBuilder builder, IReadOnlyList<OfficePoint> points, OfficeColor fill, OfficeColor? stroke = null, double strokeWidth = 0D) {
        builder.Append("<polygon")
            .AppendPointsAttribute(points)
            .AppendPaintAttribute("fill", fill);
        if (stroke.HasValue && strokeWidth > 0D) {
            builder.AppendPaintAttribute("stroke", stroke.Value)
                .AppendNumberAttribute("stroke-width", strokeWidth);
        }

        builder.Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete SVG path element using shared attribute escaping.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="pathData">SVG path data.</param>
    /// <param name="attributes">Optional already-formatted SVG attributes to append after the path data.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendPathElement(this StringBuilder builder, string pathData, string? attributes = null) {
        builder.Append("<path")
            .AppendAttribute("d", pathData);
        if (!string.IsNullOrEmpty(attributes)) {
            builder.Append(attributes);
        }

        builder.Append("/>");
        return builder;
    }

    /// <summary>
    /// Appends a complete SVG path element from shared Office path commands.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="commands">Path commands to serialize.</param>
    /// <param name="offsetX">Additional x offset applied to command points.</param>
    /// <param name="offsetY">Additional y offset applied to command points.</param>
    /// <param name="attributes">Optional already-formatted SVG attributes to append after the path data.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendPathElement(this StringBuilder builder, IReadOnlyList<OfficePathCommand> commands, double offsetX = 0D, double offsetY = 0D, string? attributes = null) {
        builder.Append("<path d=\"");
        builder.AppendPathData(commands, offsetX, offsetY)
            .Append('"');
        if (!string.IsNullOrEmpty(attributes)) {
            builder.Append(attributes);
        }

        builder.Append("/>");
        return builder;
    }

    /// <summary>
    /// Formats SVG path data for a move/line polyline or polygon using shared invariant numeric formatting.
    /// </summary>
    /// <param name="points">Points to serialize.</param>
    /// <param name="closePath">When true, appends a close-path command after at least one point.</param>
    /// <returns>SVG path data, or an empty string when no points are supplied.</returns>
    public static string FormatMoveLinePathData(IReadOnlyList<OfficePoint> points, bool closePath = false) {
        var builder = new StringBuilder();
        builder.AppendMoveLinePathData(points, closePath);
        return builder.ToString();
    }

    /// <summary>
    /// Appends SVG path data for a move/line polyline or polygon using shared invariant numeric formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="points">Points to serialize.</param>
    /// <param name="closePath">When true, appends a close-path command after at least one point.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendMoveLinePathData(this StringBuilder builder, IReadOnlyList<OfficePoint> points, bool closePath = false) {
        if (points == null) {
            throw new ArgumentNullException(nameof(points));
        }

        for (int i = 0; i < points.Count; i++) {
            builder.Append(i == 0 ? "M " : " L ");
            builder.Append(FormatNumber(points[i].X)).Append(' ').Append(FormatNumber(points[i].Y));
        }

        if (closePath && points.Count > 0) {
            builder.Append(" Z");
        }

        return builder;
    }

    /// <summary>
    /// Formats SVG path data for shared Office path commands using invariant numeric formatting.
    /// </summary>
    /// <param name="commands">Path commands to serialize.</param>
    /// <param name="offsetX">Additional x offset applied to command points.</param>
    /// <param name="offsetY">Additional y offset applied to command points.</param>
    /// <returns>SVG path data, or an empty string when no commands are supplied.</returns>
    public static string FormatPathData(IReadOnlyList<OfficePathCommand> commands, double offsetX = 0D, double offsetY = 0D) {
        var builder = new StringBuilder();
        builder.AppendPathData(commands, offsetX, offsetY);
        return builder.ToString();
    }

    /// <summary>
    /// Appends SVG path data for shared Office path commands using invariant numeric formatting.
    /// </summary>
    /// <param name="builder">Markup builder.</param>
    /// <param name="commands">Path commands to serialize.</param>
    /// <param name="offsetX">Additional x offset applied to command points.</param>
    /// <param name="offsetY">Additional y offset applied to command points.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendPathData(this StringBuilder builder, IReadOnlyList<OfficePathCommand> commands, double offsetX = 0D, double offsetY = 0D) {
        if (commands == null) {
            throw new ArgumentNullException(nameof(commands));
        }

        for (int i = 0; i < commands.Count; i++) {
            OfficePathCommand command = commands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    builder.Append('M')
                        .Append(FormatNumber(command.Point.X + offsetX)).Append(' ')
                        .Append(FormatNumber(command.Point.Y + offsetY));
                    break;
                case OfficePathCommandKind.LineTo:
                    builder.Append('L')
                        .Append(FormatNumber(command.Point.X + offsetX)).Append(' ')
                        .Append(FormatNumber(command.Point.Y + offsetY));
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    builder.Append('Q')
                        .Append(FormatNumber(command.ControlPoint1.X + offsetX)).Append(' ')
                        .Append(FormatNumber(command.ControlPoint1.Y + offsetY)).Append(' ')
                        .Append(FormatNumber(command.Point.X + offsetX)).Append(' ')
                        .Append(FormatNumber(command.Point.Y + offsetY));
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    builder.Append('C')
                        .Append(FormatNumber(command.ControlPoint1.X + offsetX)).Append(' ')
                        .Append(FormatNumber(command.ControlPoint1.Y + offsetY)).Append(' ')
                        .Append(FormatNumber(command.ControlPoint2.X + offsetX)).Append(' ')
                        .Append(FormatNumber(command.ControlPoint2.Y + offsetY)).Append(' ')
                        .Append(FormatNumber(command.Point.X + offsetX)).Append(' ')
                        .Append(FormatNumber(command.Point.Y + offsetY));
                    break;
                case OfficePathCommandKind.Close:
                    builder.Append('Z');
                    break;
            }
        }

        return builder;
    }

    /// <summary>
    /// Writes a numeric SVG attribute using shared invariant formatting.
    /// </summary>
    /// <param name="writer">XML writer receiving the attribute.</param>
    /// <param name="attributeName">Attribute name.</param>
    /// <param name="value">Numeric attribute value.</param>
    public static void WriteNumberAttribute(this XmlWriter writer, string attributeName, double value) =>
        writer.WriteAttributeString(attributeName, FormatNumber(value));

    /// <summary>
    /// Writes an SVG viewBox attribute using shared invariant numeric formatting.
    /// </summary>
    /// <param name="writer">XML writer receiving the attribute.</param>
    /// <param name="minX">Minimum x-coordinate.</param>
    /// <param name="minY">Minimum y-coordinate.</param>
    /// <param name="width">ViewBox width.</param>
    /// <param name="height">ViewBox height.</param>
    public static void WriteViewBoxAttribute(this XmlWriter writer, double minX, double minY, double width, double height) =>
        writer.WriteAttributeString("viewBox", FormatNumber(minX) + " " + FormatNumber(minY) + " " + FormatNumber(width) + " " + FormatNumber(height));

    /// <summary>
    /// Writes an SVG color attribute and matching opacity attribute when transparency is present.
    /// </summary>
    /// <param name="writer">XML writer receiving SVG attributes.</param>
    /// <param name="attributeName">SVG paint attribute name, such as <c>fill</c> or <c>stroke</c>.</param>
    /// <param name="color">Office color.</param>
    public static void WriteColorAttribute(XmlWriter writer, string attributeName, OfficeColor color) {
        writer.WriteAttributeString(attributeName, ToCssColor(color));
        if (color.A < 255) {
            writer.WriteAttributeString(attributeName + "-opacity", FormatNumber(ToOpacity(color)));
        }
    }

    /// <summary>
    /// Writes an SVG stroke-linecap attribute.
    /// </summary>
    /// <param name="writer">XML writer receiving the attribute.</param>
    /// <param name="cap">Stroke line cap.</param>
    public static void WriteStrokeLineCapAttribute(this XmlWriter writer, OfficeStrokeLineCap cap) =>
        writer.WriteAttributeString("stroke-linecap", FormatStrokeLineCap(cap));

    /// <summary>
    /// Writes an SVG stroke-linejoin attribute.
    /// </summary>
    /// <param name="writer">XML writer receiving the attribute.</param>
    /// <param name="join">Stroke line join.</param>
    public static void WriteStrokeLineJoinAttribute(this XmlWriter writer, OfficeStrokeLineJoin join) =>
        writer.WriteAttributeString("stroke-linejoin", FormatStrokeLineJoin(join));

    /// <summary>
    /// Writes an SVG stroke-dasharray attribute when a dash pattern is present.
    /// </summary>
    /// <param name="writer">XML writer receiving the attribute.</param>
    /// <param name="dashArray">SVG dash-array value, or <c>null</c> for solid strokes.</param>
    public static void WriteStrokeDashArrayAttribute(this XmlWriter writer, string? dashArray) {
        if (!string.IsNullOrEmpty(dashArray)) {
            writer.WriteAttributeString("stroke-dasharray", dashArray);
        }
    }

    /// <summary>
    /// Writes an SVG stroke-dasharray attribute for a shared Office stroke dash style.
    /// </summary>
    /// <param name="writer">XML writer receiving the attribute.</param>
    /// <param name="dashStyle">Office stroke dash style.</param>
    /// <param name="strokeWidth">Rendered stroke width.</param>
    public static void WriteStrokeDashStyleAttribute(this XmlWriter writer, OfficeStrokeDashStyle dashStyle, double strokeWidth) =>
        writer.WriteStrokeDashArrayAttribute(dashStyle.GetSvgDashArray(strokeWidth));

    /// <summary>
    /// Writes an SVG rotate transform attribute using shared invariant numeric formatting.
    /// </summary>
    /// <param name="writer">XML writer receiving the attribute.</param>
    /// <param name="degrees">Rotation angle in degrees.</param>
    /// <param name="centerX">Rotation center x-coordinate.</param>
    /// <param name="centerY">Rotation center y-coordinate.</param>
    public static void WriteRotateTransformAttribute(this XmlWriter writer, double degrees, double centerX, double centerY) =>
        writer.WriteAttributeString("transform", FormatRotateTransform(degrees, centerX, centerY));
}
