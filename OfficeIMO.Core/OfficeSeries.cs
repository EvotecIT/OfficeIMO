using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>
/// Dependency-free chart series values used by shared OfficeIMO visual renderers.
/// </summary>
public sealed partial class OfficeSeries {
    private bool _markerVisibilitySpecified = true;
    /// <summary>Largest pie or doughnut point explosion that the shared renderer supports.</summary>
    public const int MaximumPointExplosionPercent = 400;
    /// <summary>
    /// Initializes a chart series snapshot.
    /// </summary>
    /// <param name="name">Display name for the series.</param>
    /// <param name="values">Values aligned with the chart categories.</param>
    public OfficeSeries(string name, IEnumerable<double> values) : this(name, values, null, null, null, true, true) {
        _markerVisibilitySpecified = false;
    }

    /// <summary>
    /// Initializes a chart series snapshot with optional numeric X-axis values for scatter charts.
    /// </summary>
    /// <param name="name">Display name for the series.</param>
    /// <param name="values">Values aligned with the chart categories or X-axis values.</param>
    /// <param name="xValues">Optional numeric X-axis values for this series.</param>
    public OfficeSeries(string name, IEnumerable<double> values, IEnumerable<double>? xValues) : this(name, values, xValues, null, null, true, true) {
        _markerVisibilitySpecified = false;
    }

    /// <summary>
    /// Initializes a chart series snapshot with optional numeric X-axis values and source style metadata.
    /// </summary>
    /// <param name="name">Display name for the series.</param>
    /// <param name="values">Values aligned with the chart categories or X-axis values.</param>
    /// <param name="xValues">Optional numeric X-axis values for this series.</param>
    /// <param name="color">Optional source-defined series color.</param>
    public OfficeSeries(string name, IEnumerable<double> values, IEnumerable<double>? xValues, OfficeColor? color) : this(name, values, xValues, color, null, true, true) {
        _markerVisibilitySpecified = false;
    }

    /// <summary>
    /// Initializes a chart series snapshot with optional source style metadata.
    /// </summary>
    /// <param name="name">Display name for the series.</param>
    /// <param name="values">Values aligned with the chart categories or X-axis values.</param>
    /// <param name="xValues">Optional numeric X-axis values for this series.</param>
    /// <param name="color">Optional source-defined series color.</param>
    /// <param name="pointColors">Optional source-defined colors aligned with individual values.</param>
    public OfficeSeries(string name, IEnumerable<double> values, IEnumerable<double>? xValues, OfficeColor? color, IEnumerable<OfficeColor?>? pointColors) : this(name, values, xValues, color, pointColors, true, true) {
        _markerVisibilitySpecified = false;
    }

    /// <summary>
    /// Initializes a chart series snapshot with optional source style metadata and marker visibility.
    /// </summary>
    /// <param name="name">Display name for the series.</param>
    /// <param name="values">Values aligned with the chart categories or X-axis values.</param>
    /// <param name="xValues">Optional numeric X-axis values for this series.</param>
    /// <param name="color">Optional source-defined series color.</param>
    /// <param name="pointColors">Optional source-defined colors aligned with individual values.</param>
    /// <param name="showMarkers">Whether this series should render markers when the chart layout enables them.</param>
    /// <param name="showInLegend">Whether this series should appear in rendered legends.</param>
    /// <param name="connectLine">Whether this series should render connecting line segments when the chart family supports them.</param>
    /// <param name="markerSize">Optional physical marker diameter, greater than zero. Null uses the renderer or native format default.</param>
    /// <param name="markerShape">Optional source-defined marker shape.</param>
    /// <param name="markerOutlineColor">Optional source-defined marker outline color.</param>
    /// <param name="markerOutlineWidth">Optional source-defined marker outline physical width, greater than zero and at most 1584 points.</param>
    /// <param name="strokeWidth">Optional source-defined series stroke physical width, greater than zero and at most 1584 points.</param>
    /// <param name="strokeDashStyle">Optional source-defined series stroke dash style.</param>
    /// <param name="renderKind">Optional per-series chart kind used by mixed/combo chart renderers.</param>
    /// <param name="axisGroup">Primary or secondary value axis used by combo chart renderers.</param>
    public OfficeSeries(string name, IEnumerable<double> values, IEnumerable<double>? xValues, OfficeColor? color, IEnumerable<OfficeColor?>? pointColors, bool showMarkers, bool showInLegend = true, bool connectLine = true, OfficeLength? markerSize = null, OfficeChartMarkerShape? markerShape = null, OfficeColor? markerOutlineColor = null, OfficeLength? markerOutlineWidth = null, OfficeLength? strokeWidth = null, OfficeStrokeDashStyle? strokeDashStyle = null, OfficeChartType? renderKind = null, OfficeChartAxisGroup axisGroup = OfficeChartAxisGroup.Primary)
        : this(name, values, xValues, bubbleSizes: null, color, pointColors, showMarkers, showInLegend,
            connectLine, markerSize, markerShape, markerOutlineColor, markerOutlineWidth, strokeWidth,
            strokeDashStyle, renderKind, axisGroup) {
    }

    /// <summary>
    /// Creates a bubble-chart series with explicit control over whether bubble outlines are rendered.
    /// </summary>
    /// <param name="name">Display name for the series.</param>
    /// <param name="xValues">Numeric X-axis values.</param>
    /// <param name="yValues">Numeric Y-axis values.</param>
    /// <param name="bubbleSizes">Non-negative bubble sizes aligned with the X/Y values.</param>
    /// <param name="color">Optional source-defined series color.</param>
    /// <param name="pointColors">Optional source-defined colors aligned with individual bubbles.</param>
    /// <param name="showInLegend">Whether this series should appear in rendered legends.</param>
    /// <param name="markerOutlineColor">Optional source-defined bubble outline color.</param>
    /// <param name="markerOutlineWidth">Optional source-defined bubble outline physical width, greater than zero and at most 1584 points.</param>
    /// <param name="showMarkerOutline">Whether the bubble outline should be rendered.</param>
    public static OfficeSeries CreateBubble(string name, IEnumerable<double> xValues,
        IEnumerable<double> yValues, IEnumerable<double> bubbleSizes, OfficeColor? color = null,
        IEnumerable<OfficeColor?>? pointColors = null, bool showInLegend = true,
        OfficeColor? markerOutlineColor = null, OfficeLength? markerOutlineWidth = null,
        bool showMarkerOutline = true) {
        if (xValues == null) throw new ArgumentNullException(nameof(xValues));
        if (yValues == null) throw new ArgumentNullException(nameof(yValues));
        if (bubbleSizes == null) throw new ArgumentNullException(nameof(bubbleSizes));

        return new(name, yValues, xValues, bubbleSizes, color, pointColors, showMarkers: true,
            showInLegend, connectLine: false, markerSize: null, markerShape: OfficeChartMarkerShape.Circle,
            markerOutlineColor: markerOutlineColor ?? color, markerOutlineWidth, strokeWidth: null,
            strokeDashStyle: null, renderKind: OfficeChartType.Bubble,
            axisGroup: OfficeChartAxisGroup.Primary,
            showMarkerOutline: showMarkerOutline);
    }

    private OfficeSeries(string name, IEnumerable<double> values, IEnumerable<double>? xValues,
        IEnumerable<double>? bubbleSizes, OfficeColor? color, IEnumerable<OfficeColor?>? pointColors,
        bool showMarkers, bool showInLegend, bool connectLine, OfficeLength? markerSize,
        OfficeChartMarkerShape? markerShape, OfficeColor? markerOutlineColor,
        OfficeLength? markerOutlineWidth, OfficeLength? strokeWidth, OfficeStrokeDashStyle? strokeDashStyle,
        OfficeChartType? renderKind, OfficeChartAxisGroup axisGroup,
        bool showMarkerOutline = true) {
        if (values == null) {
            throw new ArgumentNullException(nameof(values));
        }
        if (markerSize.HasValue && markerSize.Value <= OfficeLength.Zero) {
            throw new ArgumentOutOfRangeException(nameof(markerSize), "Marker size must be greater than zero.");
        }
        OfficeChartStyleBounds.ValidateLineWidth(markerOutlineWidth, nameof(markerOutlineWidth));
        OfficeChartStyleBounds.ValidateLineWidth(strokeWidth, nameof(strokeWidth));

        Name = name ?? string.Empty;
        Values = new ReadOnlyCollection<double>(new List<double>(values));
        if (xValues != null) {
            XValues = new ReadOnlyCollection<double>(new List<double>(xValues));
            if (XValues.Count != Values.Count) {
                throw new ArgumentException("Series X-axis values must match the number of series values.", nameof(xValues));
            }
        }
        if (bubbleSizes != null) {
            for (int index = 0; index < Values.Count; index++) {
                double value = Values[index];
                if (double.IsNaN(value) || double.IsInfinity(value)) {
                    throw new ArgumentOutOfRangeException(nameof(values),
                        "Bubble Y-axis values must contain only finite values.");
                }
            }
            for (int index = 0; index < XValues!.Count; index++) {
                double value = XValues[index];
                if (double.IsNaN(value) || double.IsInfinity(value)) {
                    throw new ArgumentOutOfRangeException(nameof(xValues),
                        "Bubble X-axis values must contain only finite values.");
                }
            }
            BubbleSizes = new ReadOnlyCollection<double>(new List<double>(bubbleSizes));
            if (BubbleSizes.Count != Values.Count) {
                throw new ArgumentException("Series bubble sizes must match the number of series values.", nameof(bubbleSizes));
            }
            for (int index = 0; index < BubbleSizes.Count; index++) {
                double size = BubbleSizes[index];
                if (double.IsNaN(size) || double.IsInfinity(size) || size < 0D) {
                    throw new ArgumentOutOfRangeException(nameof(bubbleSizes),
                        "Bubble sizes must contain only finite, non-negative values.");
                }
            }
        }

        Color = color;
        ShowMarkers = showMarkers;
        ShowInLegend = showInLegend;
        ConnectLine = connectLine;
        MarkerSize = markerSize;
        MarkerShape = markerShape;
        MarkerOutlineColor = markerOutlineColor;
        MarkerOutlineWidth = markerOutlineWidth;
        ShowMarkerOutline = showMarkerOutline;
        StrokeWidth = strokeWidth;
        StrokeDashStyle = strokeDashStyle;
        RenderKind = renderKind;
        AxisGroup = axisGroup;
        if (pointColors != null) {
            PointColors = new ReadOnlyCollection<OfficeColor?>(new List<OfficeColor?>(pointColors));
            if (PointColors.Count != Values.Count) {
                throw new ArgumentException("Series point colors must match the number of series values.", nameof(pointColors));
            }
        }
    }

    /// <summary>Series display name.</summary>
    public string Name { get; }

    /// <summary>Series values aligned with chart categories.</summary>
    public IReadOnlyList<double> Values { get; }

    /// <summary>Optional per-series numeric X-axis values for scatter charts.</summary>
    public IReadOnlyList<double>? XValues { get; }

    /// <summary>Optional per-point sizes for bubble charts.</summary>
    public IReadOnlyList<double>? BubbleSizes { get; }

    /// <summary>Optional source-defined series color.</summary>
    public OfficeColor? Color { get; }

    /// <summary>Optional source-defined colors aligned with individual series values.</summary>
    public IReadOnlyList<OfficeColor?>? PointColors { get; }

    /// <summary>Optional series fill and outline overrides, independent of the legacy palette colour.</summary>
    public OfficeChartPointStyle? SeriesStyle { get; private set; }

    /// <summary>Returns a copy with independent series fill and outline paint. Null restores inherited paint.</summary>
    public OfficeSeries WithSeriesStyle(OfficeChartPointStyle? style) {
        var copy = CopyWithPresentation(PointStyles, PointExplosions, ShowInLegend);
        copy.SeriesStyle = style;
        return copy;
    }

    /// <summary>Optional point styles aligned with the values; explicit fills override PointColors.</summary>
    public IReadOnlyList<OfficeChartPointStyle?>? PointStyles { get; private set; }

    /// <summary>Optional pie or doughnut slice offsets, as percentages of the slice radius.</summary>
    public IReadOnlyList<int>? PointExplosions { get; private set; }

    /// <summary>
    /// Returns a copy with aligned immutable point styles. Null removes shared style metadata;
    /// null entries inherit the series/theme appearance and any legacy PointColors entry.
    /// </summary>
    public OfficeSeries WithPointStyles(IEnumerable<OfficeChartPointStyle?>? pointStyles) {
        IReadOnlyList<OfficeChartPointStyle?>? styles = pointStyles == null ? null :
            new ReadOnlyCollection<OfficeChartPointStyle?>(new List<OfficeChartPointStyle?>(pointStyles));
        if (styles != null && styles.Count != Values.Count)
            throw new ArgumentException("Point styles must match the number of series values.", nameof(pointStyles));
        return CopyWithPresentation(styles, PointExplosions, ShowInLegend);
    }

    /// <summary>Returns a copy with an explosion percentage for each pie or doughnut point. Null leaves native point offsets unchanged during data updates.</summary>
    public OfficeSeries WithPointExplosions(IEnumerable<int>? pointExplosions) {
        IReadOnlyList<int>? explosions = pointExplosions == null ? null :
            new ReadOnlyCollection<int>(new List<int>(pointExplosions));
        if (explosions != null) {
            if (explosions.Count != Values.Count)
                throw new ArgumentException("Point explosions must match the number of series values.", nameof(pointExplosions));
            foreach (int explosion in explosions)
                if (explosion < 0 || explosion > MaximumPointExplosionPercent)
                    throw new ArgumentOutOfRangeException(nameof(pointExplosions),
                        "Point explosions must be between zero and 400 percent.");
        }
        return CopyWithPresentation(PointStyles, explosions, ShowInLegend);
    }

    /// <summary>Returns a copy with the requested legend visibility and unchanged data and appearance.</summary>
    public OfficeSeries WithLegendVisibility(bool showInLegend) => CopyWithPresentation(PointStyles, PointExplosions, showInLegend);

    /// <summary>Applies the theme's line-marker policy only when the author did not choose marker visibility.</summary>
    internal bool ResolveMarkerVisibility(OfficeChartType kind, OfficeTheme theme) =>
        !HasExplicitMarkerStyle && IsLineMarkerKind(kind)
            ? theme.ChartStyle.Markers != OfficeChartMarkerShape.None || NeedsSparseLineMarkers()
            : ShowMarkers;

    internal OfficeChartMarkerShape? ResolveMarkerShape(OfficeChartType kind, OfficeTheme theme) =>
        !_markerVisibilitySpecified && IsLineMarkerKind(kind)
            ? theme.ChartStyle.Markers == OfficeChartMarkerShape.None ? OfficeChartMarkerShape.Circle : theme.ChartStyle.Markers
            : MarkerShape;

    internal bool HasExplicitMarkerVisibility => _markerVisibilitySpecified;

    // Missing values split a line into independent segments. An isolated valid value
    // needs a marker even in an otherwise long series, or it would have no painted ink.
    private bool NeedsSparseLineMarkers() {
        int finiteCount = 0;
        for (int index = 0; index < Values.Count; index++) {
            if (!IsFiniteMarkerValue(Values[index])) continue;
            finiteCount++;
            bool before = index > 0 && IsFiniteMarkerValue(Values[index - 1]);
            bool after = index + 1 < Values.Count && IsFiniteMarkerValue(Values[index + 1]);
            if (!before && !after) return true;
        }
        return finiteCount < 3;
    }

    private static bool IsFiniteMarkerValue(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    internal bool HasExplicitMarkerStyle => _markerVisibilitySpecified ||
        PointColors?.Any(color => color.HasValue) == true || PointStyles?.Any(style => style != null) == true;

    private static bool IsLineMarkerKind(OfficeChartType kind) =>
        kind is OfficeChartType.Line or OfficeChartType.LineStacked or OfficeChartType.LineStacked100;

    private OfficeSeries CopyWithPresentation(IReadOnlyList<OfficeChartPointStyle?>? styles,
        IReadOnlyList<int>? explosions, bool showInLegend) {
        return new OfficeSeries(Name, Values, XValues, BubbleSizes, Color, PointColors,
            ShowMarkers, showInLegend, ConnectLine, MarkerSize, MarkerShape, MarkerOutlineColor,
            MarkerOutlineWidth, StrokeWidth, StrokeDashStyle, RenderKind, AxisGroup, ShowMarkerOutline) {
            _markerVisibilitySpecified = this._markerVisibilitySpecified,
            SeriesStyle = SeriesStyle,
            PointStyles = styles,
            PointExplosions = explosions,
            Trendline = Trendline
        };
    }

    /// <summary>Optional trendline fitted to this series.</summary>
    public OfficeChartTrendline? Trendline { get; private set; }

    /// <summary>Returns a copy with the trendline replaced; <see langword="null"/> removes it.</summary>
    public OfficeSeries WithTrendline(OfficeChartTrendline? trendline) {
        OfficeSeries copy = CopyWithPresentation(PointStyles, PointExplosions, ShowInLegend);
        copy.Trendline = trendline;
        return copy;
    }

    /// <summary>Gets the source marker preference. Constructors without a visibility argument use the theme's
    /// line-marker defaults; explicit visibility and point styling take precedence.</summary>
    public bool ShowMarkers { get; }

    /// <summary>Optional physical marker diameter, greater than zero. Null uses the renderer or native format default.</summary>
    public OfficeLength? MarkerSize { get; }

    /// <summary>Optional source-defined marker shape.</summary>
    public OfficeChartMarkerShape? MarkerShape { get; }

    /// <summary>Optional source-defined marker outline color.</summary>
    public OfficeColor? MarkerOutlineColor { get; }

    /// <summary>Optional source-defined marker outline physical width, greater than zero and at most 1584 points.</summary>
    public OfficeLength? MarkerOutlineWidth { get; }

    /// <summary>Whether the marker or bubble outline should be rendered.</summary>
    public bool ShowMarkerOutline { get; }

    /// <summary>Optional source-defined series stroke physical width, greater than zero and at most 1584 points.</summary>
    public OfficeLength? StrokeWidth { get; }

    /// <summary>Optional source-defined series stroke dash style.</summary>
    public OfficeStrokeDashStyle? StrokeDashStyle { get; }

    /// <summary>Optional per-series chart kind for mixed/combo chart rendering.</summary>
    public OfficeChartType? RenderKind { get; }

    /// <summary>Primary or secondary value axis used by combo chart renderers.</summary>
    public OfficeChartAxisGroup AxisGroup { get; }

    /// <summary>Whether this series should appear in rendered legends.</summary>
    public bool ShowInLegend { get; }

    /// <summary>Whether this series should render connecting line segments when the chart family supports them.</summary>
    public bool ConnectLine { get; }
}
