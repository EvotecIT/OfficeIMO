using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Shared native pie and doughnut point-offset codec.</summary>
internal static class OfficeOpenXmlChartExplosions {
    internal static bool TryRead(OfficeXmlNode series, IReadOnlyList<OfficeXmlNode> points,
        int count, out int[]? explosions) {
        explosions = null;
        OfficeXmlNode seriesExplosion = series.Child(NsC, "explosion");
        if (!TryValue(seriesExplosion, out int defaultValue)) return false;
        if (seriesExplosion.Exists) {
            explosions = Enumerable.Repeat(defaultValue, count).ToArray();
        }
        var seen = new HashSet<uint>();
        foreach (OfficeXmlNode point in points) {
            OfficeXmlNode pointExplosion = point.Child(NsC, "explosion");
            if (pointExplosion.IsNone) continue;
            uint? index = point.Child(NsC, "idx").UIntVal();
            if (!index.HasValue) return false;
            // A data update can shorten the series while native point formatting remains.
            // The stale point has no visible slice to project.
            if (index.Value >= count) continue;
            if (!seen.Add(index.Value) || !TryValue(pointExplosion, out int value))
                return false;
            explosions ??= new int[count];
            explosions[(int)index.Value] = value;
        }
        return true;
    }

    internal static void ApplySeries(OfficeXmlNode series, OfficeSeries data) {
        if (data.PointExplosions == null) return;
        if (!IsPieSeries(series))
            throw new NotSupportedException("Point explosions require a native pie or doughnut series.");
        OfficeXmlNode seriesExplosion = series.Child(NsC, "explosion");
        if (seriesExplosion.Exists) seriesExplosion.Remove();
        var existing = new Dictionary<uint, OfficeXmlNode>();
        foreach (OfficeXmlNode point in series.Children(NsC, "dPt").ToList()) {
            OfficeXmlNode explosion = point.Child(NsC, "explosion");
            if (explosion.Exists) explosion.Remove();
            if (point.Child(NsC, "idx").UIntVal() is uint index && !existing.ContainsKey(index))
                existing.Add(index, point);
        }
        OfficeXmlNode anchor = FindAnchor(series);
        for (int index = 0; index < data.PointExplosions.Count; index++) {
            int percent = data.PointExplosions[index];
            if (percent == 0) continue;
            if (!existing.TryGetValue((uint)index, out OfficeXmlNode point)) {
                point = series.CreateElement(NsC, "dPt");
                point.AppendVal(NsC, "idx", Text((uint)index));
                point = series.InsertBefore(point, anchor);
                existing.Add((uint)index, point);
            }
            SetOrdered(point, point.CreateVal(NsC, "explosion", Text((uint)percent)), CDPtView.Order);
        }
    }

    internal static void ApplyPoint(OfficeXmlNode series, uint index, int? percent) {
        if (percent < 0 || percent > OfficeSeries.MaximumPointExplosionPercent)
            throw new ArgumentOutOfRangeException(nameof(percent));
        if (!IsPieSeries(series))
            throw new NotSupportedException("Point explosions require a native pie or doughnut series.");
        OfficeXmlNode point = OfficeXmlNode.None;
        foreach (OfficeXmlNode item in series.Children(NsC, "dPt")) {
            if (item.Child(NsC, "idx").UIntVal() == index) {
                point = item;
                break;
            }
        }
        int inherited = (int)(series.Child(NsC, "explosion").UIntVal() ?? 0U);
        if (!percent.HasValue || percent.Value == inherited) {
            OfficeXmlNode explosion = point.Child(NsC, "explosion");
            if (explosion.Exists) explosion.Remove();
            if (point.Exists && point.Children().ToList().TrueForAll(child => child.Is(NsC, "idx")))
                point.Remove();
            return;
        }
        if (point.IsNone) {
            point = series.CreateElement(NsC, "dPt");
            point.AppendVal(NsC, "idx", Text(index));
            point = series.InsertBefore(point, FindAnchor(series));
        }
        OfficeXmlNode previous = point.Child(NsC, "explosion");
        if (previous.Exists) previous.Remove();
        SetOrdered(point, point.CreateVal(NsC, "explosion", Text((uint)percent.Value)), CDPtView.Order);
    }

    private static OfficeXmlNode FindAnchor(OfficeXmlNode series) {
        foreach (OfficeXmlNode child in series.Children()) {
            if (child.NamespaceId == NsC && child.LocalName is "dLbls" or "cat" or "val" or "extLst") return child;
        }
        return OfficeXmlNode.None;
    }

    private static bool IsPieSeries(OfficeXmlNode series) =>
        series.Is(NsC, "ser") && series.Parent.NamespaceId == NsC &&
        series.Parent.LocalName is "pieChart" or "pie3DChart" or "doughnutChart" or "ofPieChart";

    private static bool TryValue(OfficeXmlNode source, out int value) {
        value = 0;
        if (source.IsNone) return true;
        uint? native = source.UIntVal();
        if (!native.HasValue || native.Value > OfficeSeries.MaximumPointExplosionPercent)
            return false;
        value = (int)native.Value;
        return true;
    }

    // SDK callers
    internal static bool TryRead(OpenXmlCompositeElement series, IReadOnlyList<C.DataPoint> points, int count, out int[]? explosions) =>
        TryRead(OfficeXmlNode.FromSdk(series), points.Select(point => OfficeXmlNode.FromSdk(point)).ToList(), count, out explosions);
    internal static void ApplySeries(OpenXmlCompositeElement series, OfficeSeries data) => ApplySeries(OfficeXmlNode.FromSdk(series), data);
    internal static void ApplyPoint(OpenXmlCompositeElement series, uint index, int? percent) => ApplyPoint(OfficeXmlNode.FromSdk(series), index, percent);
}
