using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Reads native series trendlines and line dash presets into the shared chart contract.</summary>
internal static class OfficeOpenXmlChartTrendlines {
    /// <summary>The largest moving-average period DrawingML allows.</summary>
    internal const int MaximumPeriod = 255;

    /// <summary>Reads the first trendline of a series, reporting what the shared renderer leaves out.</summary>
    internal static OfficeChartTrendline? Read(OpenXmlCompositeElement series, A.ColorScheme? scheme) {
        C.Trendline[] trendlines = series.Elements<C.Trendline>().Take(2).ToArray();
        if (trendlines.Length == 0) return null;
        if (trendlines.Length > 1)
            OfficeOpenXmlChartFidelityScope.Approximate("Additional trendlines on a series cannot be projected.",
                "Only the first trendline of each series is drawn.");
        C.Trendline trendline = trendlines[0];
        OfficeChartTrendlineType? type = trendline.GetFirstChild<C.TrendlineType>()?.Val?.InnerText switch {
            "exp" => OfficeChartTrendlineType.Exponential,
            "linear" => OfficeChartTrendlineType.Linear,
            "log" => OfficeChartTrendlineType.Logarithmic,
            "movingAvg" => OfficeChartTrendlineType.MovingAverage,
            "poly" => OfficeChartTrendlineType.Polynomial,
            "power" => OfficeChartTrendlineType.Power,
            _ => null
        };
        int? order = trendline.GetFirstChild<C.PolynomialOrder>()?.Val?.Value is byte polynomialOrder ? polynomialOrder : null;
        int? period = trendline.GetFirstChild<C.Period>()?.Val?.Value is uint averagePeriod && averagePeriod <= int.MaxValue ? (int)averagePeriod : null;
        if (type == null || type == OfficeChartTrendlineType.Polynomial && order is null or < 2 or > 6 ||
            type == OfficeChartTrendlineType.MovingAverage && period is null or < 2 or > MaximumPeriod) {
            OfficeOpenXmlChartFidelityScope.Approximate("The native trendline type cannot be projected.",
                "A trendline the renderer cannot fit is left out.");
            return null;
        }
        if (series.Parent is C.BarChart bars && bars.BarDirection?.Val?.Value == C.BarDirectionValues.Bar) {
            OfficeOpenXmlChartFidelityScope.Approximate("Trendlines on horizontal bar charts cannot be projected.",
                "Trendlines on horizontal bars are left out.");
            return null;
        }
        if (trendline.GetFirstChild<C.Forward>()?.Val?.Value is double forward && forward != 0D ||
            trendline.GetFirstChild<C.Backward>()?.Val?.Value is double backward && backward != 0D)
            OfficeOpenXmlChartFidelityScope.Approximate("Trendline forecasts cannot be projected.",
                "Trendline forecasts are not extended beyond the data.");
        if (trendline.GetFirstChild<C.TrendlineLabel>() is C.TrendlineLabel label &&
            (label.GetFirstChild<C.ChartText>() != null || label.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>() != null))
            OfficeOpenXmlChartFidelityScope.Approximate("Custom trendline labels cannot be projected.",
                "A trendline label with its own text or position is drawn in the standard place.");
        if (HasLegendEntry(series))
            OfficeOpenXmlChartFidelityScope.Approximate("Trendline legend entries cannot be projected.",
                "Trendline entries are left out of the legend.");
        C.ChartShapeProperties? properties = trendline.GetFirstChild<C.ChartShapeProperties>();
        A.Outline? outline = properties?.GetFirstChild<A.Outline>();
        OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(outline?.GetFirstChild<A.SolidFill>(), scheme);
        OfficeLength? width = outline?.Width?.Value is int emus && emus > 0 && emus / 12700D <= OfficeChartStyleBounds.MaximumLineWidthPoints
            ? OfficeLength.Emus(emus) : null;
        double? intercept = trendline.GetFirstChild<C.Intercept>()?.Val?.Value is double fixedIntercept &&
            !double.IsNaN(fixedIntercept) && !double.IsInfinity(fixedIntercept) ? fixedIntercept : null;
        bool hidden = HidesLine(outline);
        var result = new OfficeChartTrendline(type.Value, order, period,
            displayEquation: trendline.GetFirstChild<C.DisplayEquation>()?.Val?.Value == true,
            displayRSquared: trendline.GetFirstChild<C.DisplayRSquaredValue>()?.Val?.Value == true,
            color: color, width: width, intercept: intercept);
        return hidden ? result.WithHiddenLine() : result.WithDashStyle(ReadDashOrApproximate(outline));
    }

    /// <summary>True when the trendline's outline has no fill, so only its labels show.</summary>
    internal static bool HidesLine(A.Outline? outline) => outline?.GetFirstChild<A.NoFill>() != null;

    // Office lists each trendline in the legend after every series, unless that entry is deleted. Only the chart
    // groups and their series are visited, never the cached points.
    private static bool HasLegendEntry(OpenXmlCompositeElement series) {
        C.Chart? chart = series.Ancestors<C.Chart>().FirstOrDefault();
        C.Legend? legend = chart?.GetFirstChild<C.Legend>();
        if (legend == null || chart?.PlotArea == null) return false;
        int seriesCount = 0, ordinal = 0;
        bool found = false;
        foreach (OpenXmlElement layer in chart.PlotArea.ChildElements) {
            if (!layer.LocalName.EndsWith("Chart", StringComparison.Ordinal)) continue;
            foreach (OpenXmlElement item in layer.ChildElements) {
                if (item.LocalName != "ser") continue;
                seriesCount++;
                if (ReferenceEquals(item, series)) found = true;
                else if (!found && item.ChildElements.OfType<C.Trendline>().Any()) ordinal++;
            }
        }
        uint entry = (uint)(seriesCount + ordinal);
        return !legend.Elements<C.LegendEntry>().Any(item => item.GetFirstChild<C.Index>()?.Val?.Value == entry &&
            item.GetFirstChild<C.Delete>() is C.Delete deleted && deleted.Val?.Value != false);
    }

    /// <summary>Reads the trendline dash, reporting a preset the renderer cannot draw.</summary>
    internal static OfficeStrokeDashStyle? ReadDashOrApproximate(A.Outline? outline) {
        OfficeStrokeDashStyle? dash = ReadDash(outline);
        if ((outline?.GetFirstChild<A.PresetDash>() != null || outline?.GetFirstChild<A.CustomDash>() != null) && !dash.HasValue)
            OfficeOpenXmlChartFidelityScope.Approximate("The trendline dash pattern cannot be projected.",
                "A trendline dash pattern the renderer cannot draw is drawn solid.");
        return dash;
    }

    /// <summary>
    /// Maps a DrawingML preset dash to the shared pattern. The <c>sys</c> and long presets differ from their
    /// base pattern only in segment length relative to the line width; best-effort projections draw them with
    /// the nearest pattern and report it.
    /// </summary>
    internal static OfficeStrokeDashStyle? ReadDash(A.Outline? outline) {
        string? preset = outline?.GetFirstChild<A.PresetDash>()?.Val?.InnerText;
        switch (preset) {
            case null: return null;
            case "solid": return OfficeStrokeDashStyle.Solid;
            case "dash": return OfficeStrokeDashStyle.Dash;
            case "dot": return OfficeStrokeDashStyle.Dot;
            case "dashDot": return OfficeStrokeDashStyle.DashDot;
            case "lgDashDotDot": return OfficeStrokeDashStyle.DashDotDot;
        }
        OfficeStrokeDashStyle? nearest = preset switch {
            "sysDash" or "lgDash" => OfficeStrokeDashStyle.Dash,
            "sysDot" => OfficeStrokeDashStyle.Dot,
            "sysDashDot" or "lgDashDot" => OfficeStrokeDashStyle.DashDot,
            "sysDashDotDot" => OfficeStrokeDashStyle.DashDotDot,
            _ => null
        };
        if (!nearest.HasValue || !OfficeOpenXmlChartFidelityScope.AllowsApproximation) return null;
        OfficeOpenXmlChartFidelityScope.Record("The '" + preset + "' dash pattern is drawn with the nearest shared dash pattern.");
        return nearest;
    }
}
