using System.Collections.Generic;
using OfficeIMO;

namespace OfficeIMO.Tests;

/// <summary>Word's shared chart composition contract, independent of other formats' chart families.</summary>
internal static class WordSharedChartCases {
    internal static IReadOnlyList<OfficeChartType> Kinds { get; } = new[] {
        OfficeChartType.ColumnClustered, OfficeChartType.ColumnStacked, OfficeChartType.ColumnStacked100,
        OfficeChartType.BarClustered, OfficeChartType.BarStacked, OfficeChartType.BarStacked100,
        OfficeChartType.Line, OfficeChartType.LineStacked, OfficeChartType.LineStacked100,
        OfficeChartType.Area, OfficeChartType.AreaStacked, OfficeChartType.AreaStacked100,
        OfficeChartType.Scatter, OfficeChartType.Radar, OfficeChartType.Pie, OfficeChartType.Doughnut, OfficeChartType.Bubble
    };
}
