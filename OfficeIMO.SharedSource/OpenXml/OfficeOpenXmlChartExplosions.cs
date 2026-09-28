using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Shared native pie and doughnut point-offset codec.</summary>
internal static class OfficeOpenXmlChartExplosions {
    internal static bool TryRead(OpenXmlCompositeElement series, IReadOnlyList<C.DataPoint> points,
        int count, out int[]? explosions) {
        explosions = null;
        C.Explosion? seriesExplosion = series.GetFirstChild<C.Explosion>();
        if (!TryValue(seriesExplosion, out int defaultValue)) return false;
        if (seriesExplosion != null) {
            explosions = Enumerable.Repeat(defaultValue, count).ToArray();
        }
        var seen = new HashSet<uint>();
        foreach (C.DataPoint point in points) {
            C.Explosion? pointExplosion = point.GetFirstChild<C.Explosion>();
            if (pointExplosion == null) continue;
            uint? index = point.Index?.Val?.Value;
            if (!index.HasValue || index.Value >= count || !seen.Add(index.Value) ||
                !TryValue(pointExplosion, out int value))
                return false;
            explosions ??= new int[count];
            explosions[(int)index.Value] = value;
        }
        return true;
    }

    internal static void ApplySeries(OpenXmlCompositeElement series, OfficeChartSeries data) {
        if (data.PointExplosions == null) return;
        if (series is not C.PieChartSeries)
            throw new NotSupportedException("Point explosions require a native pie or doughnut series.");
        series.GetFirstChild<C.Explosion>()?.Remove();
        foreach (C.DataPoint point in series.Elements<C.DataPoint>())
            point.GetFirstChild<C.Explosion>()?.Remove();
        for (int index = 0; index < data.PointExplosions.Count; index++) {
            int percent = data.PointExplosions[index];
            if (percent == 0) continue;
            C.DataPoint? point = series.Elements<C.DataPoint>()
                .FirstOrDefault(item => item.Index?.Val?.Value == (uint)index);
            if (point == null) {
                point = new C.DataPoint(new C.Index { Val = (uint)index });
                OpenXmlElement? anchor = series.ChildElements.FirstOrDefault(child =>
                    child is C.DataLabels or C.CategoryAxisData or C.Values or C.ExtensionList);
                if (anchor != null) series.InsertBefore(point, anchor);
                else series.Append(point);
            }
            point.AddChild(new C.Explosion { Val = (uint)percent }, true);
        }
    }

    private static bool TryValue(C.Explosion? source, out int value) {
        value = 0;
        if (source == null) return true;
        uint? native = source.Val?.Value;
        if (!native.HasValue || native.Value > OfficeChartSeries.MaximumPointExplosionPercent)
            return false;
        value = (int)native.Value;
        return true;
    }
}
