using System.Globalization;
using System.Text;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Radial geometry metadata shared by semantic Office HTML adapters.</summary>
internal static class OfficeHtmlChartRadialLayout {
    internal static void AppendAttributes(StringBuilder html, OfficeChartRadialLayout layout) {
        html.Append(" data-officeimo-first-slice-angle=\"")
            .Append(layout.FirstSliceAngleDegrees.ToString(CultureInfo.InvariantCulture))
            .Append("\" data-officeimo-doughnut-hole=\"")
            .Append(layout.DoughnutHolePercent.ToString(CultureInfo.InvariantCulture)).Append('"');
    }

    internal static bool TryRead(IElement item, out OfficeChartRadialLayout layout) {
        layout = OfficeChartRadialLayout.Default;
        IElement? table = item.QuerySelector("table.officeimo-chart-data");
        string? rawAngle = table?.GetAttribute("data-officeimo-first-slice-angle");
        string? rawHole = table?.GetAttribute("data-officeimo-doughnut-hole");
        int angle = 0, hole = 50;
        if (rawAngle != null && (!int.TryParse(rawAngle, NumberStyles.None, CultureInfo.InvariantCulture, out angle) || angle > 360)) return false;
        if (rawHole != null && (!int.TryParse(rawHole, NumberStyles.None, CultureInfo.InvariantCulture, out hole) || hole < 10 || hole > 90)) return false;
        layout = new OfficeChartRadialLayout(angle, hole);
        return true;
    }
}
