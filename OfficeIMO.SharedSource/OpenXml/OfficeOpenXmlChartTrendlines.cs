using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Reads native series trendlines and line dash presets into the shared chart contract.</summary>
internal static class OfficeOpenXmlChartTrendlines {
    /// <summary>The largest moving-average period DrawingML allows.</summary>
    internal const int MaximumPeriod = 255;

    /// <summary>Reads the first trendline of a series, reporting what the shared renderer leaves out.</summary>
    internal static OfficeChartTrendline? Read(OfficeXmlNode series, OfficeXmlNode scheme, string seriesName) {
        OfficeXmlNode trendline = series.Child(NsC, "trendline");
        if (trendline.IsNone) return null;
        if (trendline.NextSiblingNamed(NsC, "trendline").Exists)
            OfficeOpenXmlChartFidelityScope.Approximate("Additional trendlines on a series cannot be projected.",
                "Only the first trendline of each series is drawn.");
        OfficeChartTrendlineType? type = trendline.ChildVal(NsC, "trendlineType") switch {
            "exp" => OfficeChartTrendlineType.Exponential,
            "linear" => OfficeChartTrendlineType.Linear,
            "log" => OfficeChartTrendlineType.Logarithmic,
            "movingAvg" => OfficeChartTrendlineType.MovingAverage,
            "poly" => OfficeChartTrendlineType.Polynomial,
            "power" => OfficeChartTrendlineType.Power,
            _ => null
        };
        int? order = trendline.Child(NsC, "order").ByteVal() is byte polynomialOrder ? polynomialOrder : null;
        int? period = trendline.Child(NsC, "period").UIntVal() is uint averagePeriod && averagePeriod <= int.MaxValue ? (int)averagePeriod : null;
        if (type == null || type == OfficeChartTrendlineType.Polynomial && order is null or < 2 or > 6 ||
            type == OfficeChartTrendlineType.MovingAverage && period is null or < 2 or > MaximumPeriod) {
            OfficeOpenXmlChartFidelityScope.Approximate("The native trendline type cannot be projected.",
                "A trendline the renderer cannot fit is left out.");
            return null;
        }
        OfficeXmlNode label = trendline.Child(NsC, "trendlineLbl");
        if (label.Exists &&
            (label.Child(NsC, "tx").Exists || label.Child(NsC, "layout").Child(NsC, "manualLayout").Exists))
            OfficeOpenXmlChartFidelityScope.Approximate("Custom trendline labels cannot be projected.",
                "A trendline label with its own text or position is drawn in the standard place.");
        OfficeXmlNode outline = trendline.Child(NsC, "spPr").Child(NsA, "ln");
        OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(outline.Child(NsA, "solidFill"), scheme);
        OfficeLength? width = outline.IntAttr("w") is int emus && emus > 0 && emus / 12700D <= OfficeChartStyleBounds.MaximumLineWidthPoints
            ? OfficeLength.Emus(emus) : null;
        double? intercept = trendline.Child(NsC, "intercept").DoubleVal() is double fixedIntercept &&
            !double.IsNaN(fixedIntercept) && !double.IsInfinity(fixedIntercept) ? fixedIntercept : null;
        bool hidden = HidesLine(outline);
        var result = new OfficeChartTrendline(type.Value, order, period,
            displayEquation: trendline.Child(NsC, "dispEq").BoolVal() == true,
            displayRSquared: trendline.Child(NsC, "dispRSqr").BoolVal() == true,
            color: color, width: width, intercept: intercept)
            .WithNativeDetails(trendline.Child(NsC, "forward").DoubleVal() ?? 0D, trendline.Child(NsC, "backward").DoubleVal() ?? 0D,
                ReadLegendName(series, trendline, type.Value, order, period, seriesName));
        return hidden ? result.WithHiddenLine() : result.WithDashStyle(ReadDashOrApproximate(outline));
    }

    /// <summary>True when the trendline's outline has no fill, so only its labels show.</summary>
    internal static bool HidesLine(OfficeXmlNode outline) => outline.Child(NsA, "noFill").Exists;

    // Office lists each trendline in the legend after every series, unless that entry is deleted. Only the chart
    // groups and their series are visited, never the cached points.
    private static bool HasLegendEntry(OfficeXmlNode series) {
        OfficeXmlNode chart = series.Ancestor(NsC, "chart");
        OfficeXmlNode legend = chart.Child(NsC, "legend");
        OfficeXmlNode plotArea = chart.Child(NsC, "plotArea");
        if (legend.IsNone || plotArea.IsNone) return false;
        int seriesCount = 0, ordinal = 0;
        bool found = false;
        foreach (OfficeXmlNode layer in plotArea.Children()) {
            if (!layer.LocalName.EndsWith("Chart", StringComparison.Ordinal)) continue;
            foreach (OfficeXmlNode item in layer.Children()) {
                if (item.LocalName != "ser") continue;
                seriesCount++;
                if (item == series) found = true;
                else if (!found && item.Child(NsC, "trendline").Exists) ordinal++;
            }
        }
        uint entry = (uint)(seriesCount + ordinal);
        foreach (OfficeXmlNode item in legend.Children(NsC, "legendEntry")) {
            if (item.Child(NsC, "idx").UIntVal() != entry) continue;
            OfficeXmlNode deleted = item.Child(NsC, "delete");
            if (deleted.Exists && deleted.BoolVal() != false) return false;
        }
        return true;
    }

    /// <summary>Reads the trendline dash, reporting a preset the renderer cannot draw.</summary>
    internal static OfficeStrokeDashStyle? ReadDashOrApproximate(OfficeXmlNode outline) {
        OfficeStrokeDashStyle? dash = ReadDash(outline);
        if ((outline.Child(NsA, "prstDash").Exists || outline.Child(NsA, "custDash").Exists) && !dash.HasValue)
            OfficeOpenXmlChartFidelityScope.Approximate("The trendline dash pattern cannot be projected.",
                "A trendline dash pattern the renderer cannot draw is drawn solid.");
        return dash;
    }

    // Office lists each trendline in the legend after the series, under its own name or an automatic one,
    // unless that legend entry is deleted.
    internal static string? ReadLegendName(OfficeXmlNode series, OfficeXmlNode trendline, OfficeChartTrendlineType type,
        int? order, int? period, string seriesName) {
        if (!HasLegendEntry(series)) return null;
        string? own = trendline.ChildText(NsC, "name");
        return string.IsNullOrWhiteSpace(own)
            ? OfficeChartDrawingRenderer.GetTrendlineLegendName(new OfficeChartTrendline(type, order, period), seriesName)
            : own;
    }

    /// <summary>
    /// Maps a DrawingML preset dash to the shared pattern. The <c>sys</c> and long presets differ from their
    /// base pattern only in segment length relative to the line width; best-effort projections draw them with
    /// the nearest pattern and report it.
    /// </summary>
    internal static OfficeStrokeDashStyle? ReadDash(OfficeXmlNode outline) {
        string? preset = outline.ChildVal(NsA, "prstDash");
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

    // SDK callers
    internal static OfficeChartTrendline? Read(OpenXmlCompositeElement series, A.ColorScheme? scheme, string seriesName) => Read(OfficeXmlNode.FromSdk(series), OfficeXmlNode.FromSdk(scheme), seriesName);
    internal static bool HidesLine(A.Outline? outline) => HidesLine(OfficeXmlNode.FromSdk(outline));
    internal static OfficeStrokeDashStyle? ReadDashOrApproximate(A.Outline? outline) => ReadDashOrApproximate(OfficeXmlNode.FromSdk(outline));
    internal static string? ReadLegendName(OpenXmlCompositeElement series, C.Trendline trendline, OfficeChartTrendlineType type, int? order, int? period, string seriesName) =>
        ReadLegendName(OfficeXmlNode.FromSdk(series), OfficeXmlNode.FromSdk(trendline), type, order, period, seriesName);
    internal static OfficeStrokeDashStyle? ReadDash(A.Outline? outline) => ReadDash(OfficeXmlNode.FromSdk(outline));
}
