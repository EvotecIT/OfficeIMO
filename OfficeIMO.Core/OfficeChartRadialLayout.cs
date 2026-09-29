using System;

namespace OfficeIMO.Drawing;

/// <summary>Native-compatible geometry for pie and doughnut charts.</summary>
public sealed class OfficeChartRadialLayout {
    /// <summary>Creates radial geometry with a clockwise rotation and doughnut hole percentage.</summary>
    /// <param name="firstSliceAngleDegrees">Clockwise rotation from the top, from zero through 360 degrees.</param>
    /// <param name="doughnutHolePercent">Inner diameter as a percentage of outer diameter, from 10 through 90.</param>
    public OfficeChartRadialLayout(int firstSliceAngleDegrees = 0, int doughnutHolePercent = 50) {
        if (firstSliceAngleDegrees < 0 || firstSliceAngleDegrees > 360)
            throw new ArgumentOutOfRangeException(nameof(firstSliceAngleDegrees));
        if (doughnutHolePercent < 10 || doughnutHolePercent > 90)
            throw new ArgumentOutOfRangeException(nameof(doughnutHolePercent));
        FirstSliceAngleDegrees = firstSliceAngleDegrees;
        DoughnutHolePercent = doughnutHolePercent;
    }

    /// <summary>Clockwise rotation of the first slice boundary from the top.</summary>
    public int FirstSliceAngleDegrees { get; }
    /// <summary>Inner diameter as a percentage of the doughnut outer diameter.</summary>
    public int DoughnutHolePercent { get; }
    /// <summary>Default native geometry: first boundary at the top and a 50-percent hole.</summary>
    public static OfficeChartRadialLayout Default { get; } = new();
}
