using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        private static void PreserveSharedChartFormatting(
            C.PlotArea source, C.PlotArea replacement,
            ISet<uint> preservedSeriesIndexes) {
            var axisBindings = new Dictionary<uint, uint>();
            PreserveSharedChartLayers(
                source, replacement, preservedSeriesIndexes, axisBindings);
            PreserveSharedAxes(source, replacement, axisBindings);
        }

        private static void PreserveSharedChartLayers(
            C.PlotArea source, C.PlotArea replacement,
            ISet<uint> preservedSeriesIndexes, IDictionary<uint, uint> axisBindings) {
            List<OpenXmlCompositeElement> sourceLayers = source.ChildElements
                .OfType<OpenXmlCompositeElement>().Where(IsSharedChartLayer).ToList();
            var usedLayers = new HashSet<OpenXmlCompositeElement>();
            foreach (OpenXmlCompositeElement generated in replacement.ChildElements
                         .OfType<OpenXmlCompositeElement>().Where(IsSharedChartLayer).ToList()) {
                List<OpenXmlCompositeElement> matches = sourceLayers.Where(candidate =>
                    !usedLayers.Contains(candidate) &&
                    AreCompatibleSharedChartLayers(candidate, generated, source, replacement)).ToList();
                if (matches.Count == 0) continue;
                List<OpenXmlCompositeElement> generatedSeries = generated.ChildElements.OfType<OpenXmlCompositeElement>()
                    .Where(IsSharedSeriesElement).ToList();
                int offset = 0;
                for (int layerIndex = 0; layerIndex < matches.Count && offset < generatedSeries.Count; layerIndex++) {
                    OpenXmlCompositeElement match = matches[layerIndex];
                    int remaining = generatedSeries.Count - offset;
                    int oldCount = Math.Max(1, match.ChildElements.OfType<OpenXmlCompositeElement>().Count(IsSharedSeriesElement));
                    int count = layerIndex == matches.Count - 1 ? remaining : Math.Min(oldCount, remaining);
                    var slice = (OpenXmlCompositeElement)generated.CloneNode(true);
                    foreach (OpenXmlCompositeElement item in slice.ChildElements.OfType<OpenXmlCompositeElement>().Where(IsSharedSeriesElement).ToList()) item.Remove();
                    foreach (OpenXmlCompositeElement item in generatedSeries.Skip(offset).Take(count)) InsertSeries(slice, item.CloneNode(true));
                    var preserved = (OpenXmlCompositeElement)match.CloneNode(true);
                    ReplaceSharedSeriesData(preserved, slice, preservedSeriesIndexes);
                    BindSharedAxisReferences(match, generated, source, replacement, axisBindings);
                    ReplaceSharedAxisReferences(preserved, generated);
                    replacement.InsertBefore(preserved, generated);
                    usedLayers.Add(match);
                    offset += count;
                }
                generated.Remove();
            }
        }

        private static bool IsSharedChartLayer(OpenXmlCompositeElement element) =>
            element is C.BarChart || element is C.LineChart || element is C.AreaChart ||
            element is C.RadarChart || element is C.PieChart || element is C.DoughnutChart ||
            element is C.BubbleChart;

        private static bool AreCompatibleSharedChartLayers(OpenXmlCompositeElement source,
            OpenXmlCompositeElement replacement, C.PlotArea sourcePlotArea, C.PlotArea replacementPlotArea) {
            if (source.GetType() != replacement.GetType()) return false;
            if (source is C.BarChart sourceBar && replacement is C.BarChart replacementBar) {
                if (sourceBar.BarDirection?.Val?.Value != replacementBar.BarDirection?.Val?.Value ||
                    sourceBar.BarGrouping?.Val?.Value != replacementBar.BarGrouping?.Val?.Value) return false;
            } else if (source is C.LineChart sourceLine && replacement is C.LineChart replacementLine) {
                if (sourceLine.Grouping?.Val?.Value != replacementLine.Grouping?.Val?.Value) return false;
            } else if (source is C.AreaChart sourceArea && replacement is C.AreaChart replacementArea) {
                if (sourceArea.Grouping?.Val?.Value != replacementArea.Grouping?.Val?.Value) return false;
            }

            return IsSecondarySharedChartLayer(source, sourcePlotArea) ==
                   IsSecondarySharedChartLayer(replacement, replacementPlotArea);
        }

        private static bool IsSecondarySharedChartLayer(OpenXmlCompositeElement chartLayer,
            C.PlotArea plotArea) {
            return OfficeOpenXmlChartAxisGroups.Read(plotArea, chartLayer) == OfficeIMO.Drawing.OfficeChartAxisGroup.Secondary;
        }

        private static void ReplaceSharedSeriesData(
            OpenXmlCompositeElement preserved,
            OpenXmlCompositeElement generated,
            ISet<uint> preservedSeriesIndexes) {
            List<OpenXmlCompositeElement> oldSeries = preserved.ChildElements
                .OfType<OpenXmlCompositeElement>().Where(IsSharedSeriesElement).ToList();
            var usedSeries = new HashSet<OpenXmlCompositeElement>();
            OpenXmlElement? insertionPoint = oldSeries.FirstOrDefault();
            List<OpenXmlCompositeElement> generatedSeriesElements = generated.ChildElements
                .OfType<OpenXmlCompositeElement>().Where(IsSharedSeriesElement).ToList();
            for (int position = 0; position < generatedSeriesElements.Count; position++) {
                OpenXmlCompositeElement generatedSeries =
                    generatedSeriesElements[position];
                uint? seriesIndex = generatedSeries.GetFirstChild<C.Index>()?.Val?.Value;
                OpenXmlCompositeElement? sourceSeries = oldSeries.FirstOrDefault(series =>
                    !usedSeries.Contains(series) &&
                    series.GetType() == generatedSeries.GetType() &&
                    series.GetFirstChild<C.Index>()?.Val?.Value == seriesIndex);
                if (sourceSeries == null && position < oldSeries.Count &&
                    !usedSeries.Contains(oldSeries[position]) &&
                    oldSeries[position].GetType() == generatedSeries.GetType()) {
                    sourceSeries = oldSeries[position];
                }
                OpenXmlCompositeElement updated = sourceSeries == null
                    ? (OpenXmlCompositeElement)generatedSeries.CloneNode(true)
                    : UpdateSharedSeriesData(sourceSeries, generatedSeries);
                if (sourceSeries != null) {
                    usedSeries.Add(sourceSeries);
                    if (seriesIndex.HasValue) {
                        preservedSeriesIndexes.Add(seriesIndex.Value);
                    }
                }
                if (insertionPoint == null) preserved.AddChild(updated, true);
                else preserved.InsertBefore(updated, insertionPoint);
            }
            foreach (OpenXmlCompositeElement series in oldSeries) series.Remove();
        }

        private static OpenXmlCompositeElement UpdateSharedSeriesData(OpenXmlCompositeElement source,
            OpenXmlCompositeElement generated) {
            var updated = (OpenXmlCompositeElement)source.CloneNode(true);
            ReplaceSharedSeriesChild<C.Index>(updated, generated);
            ReplaceSharedSeriesChild<C.Order>(updated, generated);
            ReplaceSharedSeriesChild<C.SeriesText>(updated, generated);
            ReplaceSharedSeriesChild<C.CategoryAxisData>(updated, generated);
            ReplaceSharedSeriesChild<C.Values>(updated, generated);
            ReplaceSharedSeriesChild<C.XValues>(updated, generated);
            ReplaceSharedSeriesChild<C.YValues>(updated, generated);
            ReplaceSharedSeriesChild<C.BubbleSize>(updated, generated);
            return updated;
        }

        private static void ReplaceSharedSeriesChild<T>(OpenXmlCompositeElement updated,
            OpenXmlCompositeElement generated) where T : OpenXmlElement {
            T? current = updated.GetFirstChild<T>();
            T? replacement = generated.GetFirstChild<T>();
            if (replacement == null) {
                current?.Remove();
            } else if (current == null) {
                updated.AddChild(replacement.CloneNode(true), true);
            } else {
                updated.ReplaceChild(replacement.CloneNode(true), current);
            }
        }

        private static bool IsSharedSeriesElement(OpenXmlCompositeElement element) =>
            element is C.BarChartSeries || element is C.LineChartSeries ||
            element is C.AreaChartSeries || element is C.RadarChartSeries ||
            element is C.PieChartSeries || element is C.BubbleChartSeries;

        private static void ReplaceSharedAxisReferences(OpenXmlCompositeElement preserved,
            OpenXmlCompositeElement generated) {
            List<C.AxisId> preservedIds = preserved.Elements<C.AxisId>().ToList();
            List<C.AxisId> generatedIds = generated.Elements<C.AxisId>().ToList();
            if (preservedIds.Count != generatedIds.Count) return;
            for (int index = 0; index < preservedIds.Count; index++) {
                preservedIds[index].Val = generatedIds[index].Val;
            }
        }

        private static void BindSharedAxisReferences(OpenXmlCompositeElement source,
            OpenXmlCompositeElement generated, C.PlotArea sourcePlot, C.PlotArea generatedPlot,
            IDictionary<uint, uint> bindings) {
            List<C.AxisId> sourceIds = source.Elements<C.AxisId>().ToList();
            List<C.AxisId> generatedIds = generated.Elements<C.AxisId>().ToList();
            if (sourceIds.Count != generatedIds.Count) return;
            for (int index = 0; index < sourceIds.Count; index++) {
                uint? generatedId = generatedIds[index].Val?.Value;
                uint? sourceId = sourceIds[index].Val?.Value;
                if (!(source is C.BubbleChart) && !(source is C.ScatterChart)) {
                    OpenXmlCompositeElement? generatedAxis = generatedPlot.ChildElements.OfType<OpenXmlCompositeElement>()
                        .SingleOrDefault(axis => (axis is C.ValueAxis || IsSharedCategoryAxis(axis)) &&
                            axis.GetFirstChild<C.AxisId>()?.Val?.Value == generatedId);
                    if (generatedAxis == null) throw new NotSupportedException("A chart axis reference must resolve to a unique category or value axis.");
                    var referencedIds = new HashSet<uint>(sourceIds.Where(id => id.Val != null).Select(id => id.Val!.Value));
                    var matches = sourcePlot.ChildElements.OfType<OpenXmlCompositeElement>().Where(axis =>
                        axis.GetFirstChild<C.AxisId>()?.Val is { } id && referencedIds.Contains(id.Value) &&
                        (generatedAxis is C.ValueAxis ? axis is C.ValueAxis : IsSharedCategoryAxis(axis))).Take(2).ToList();
                    if (matches.Count != 1) throw new NotSupportedException("A category chart layer must reference one category axis and one value axis.");
                    sourceId = matches[0].GetFirstChild<C.AxisId>()!.Val!.Value;
                }
                if (!sourceId.HasValue || !generatedId.HasValue) continue;
                if (bindings.TryGetValue(generatedId.Value, out uint previous) && previous != sourceId.Value)
                    throw new NotSupportedException("Native layers in the same chart family and axis group must share their axis references for a shared data update.");
                bindings[generatedId.Value] = sourceId.Value;
            }
        }

        private static bool IsSharedCategoryAxis(OpenXmlCompositeElement axis) =>
            axis is C.CategoryAxis || axis is C.DateAxis;

        private static bool UsesHorizontalSharedAxes(C.PlotArea plotArea) =>
            plotArea.Elements<C.BarChart>().Any(chart =>
                chart.BarDirection?.Val?.Value == C.BarDirectionValues.Bar);

        private static void PreserveSharedAxes(C.PlotArea source, C.PlotArea replacement,
            IReadOnlyDictionary<uint, uint> bindings) {
            if (UsesHorizontalSharedAxes(source) != UsesHorizontalSharedAxes(replacement)) return;
            List<OpenXmlCompositeElement> sourceAxes = source.ChildElements.OfType<OpenXmlCompositeElement>()
                .Where(axis => axis is C.ValueAxis || IsSharedCategoryAxis(axis)).ToList();
            List<OpenXmlCompositeElement> generatedAxes = replacement.ChildElements.OfType<OpenXmlCompositeElement>()
                .Where(axis => axis is C.ValueAxis || IsSharedCategoryAxis(axis)).ToList();
            var used = new HashSet<OpenXmlCompositeElement>();
            foreach (OpenXmlCompositeElement generated in generatedAxes) {
                bool Compatible(OpenXmlCompositeElement axis) => !used.Contains(axis) &&
                    (axis.GetType() == generated.GetType() || IsSharedCategoryAxis(axis) && IsSharedCategoryAxis(generated));
                OpenXmlCompositeElement? match = null;
                uint? generatedId = generated.GetFirstChild<C.AxisId>()?.Val?.Value;
                if (generatedId.HasValue && bindings.TryGetValue(generatedId.Value, out uint sourceId)) {
                    match = sourceAxes.FirstOrDefault(axis => Compatible(axis) && axis.GetFirstChild<C.AxisId>()?.Val?.Value == sourceId);
                } else {
                    // A family change can still retain an unambiguous axis in the same
                    // orientation. XML element order never determines axis identity.
                    var candidates = sourceAxes.Where(axis => Compatible(axis) &&
                        axis.GetFirstChild<C.AxisPosition>()?.Val?.Value == generated.GetFirstChild<C.AxisPosition>()?.Val?.Value).Take(2).ToList();
                    if (candidates.Count == 1) match = candidates[0];
                }
                if (match == null) continue;
                used.Add(match);
                ReplaceSharedAxis(replacement, match, generated);
            }
        }
        private static void ReplaceSharedAxis(C.PlotArea replacement, OpenXmlCompositeElement source,
            OpenXmlCompositeElement generated) {
            var preserved = (OpenXmlCompositeElement)source.CloneNode(true);
            C.AxisId? generatedId = generated.GetFirstChild<C.AxisId>();
            C.CrossingAxis? generatedCrossing = generated.GetFirstChild<C.CrossingAxis>();
            if (generatedId != null) {
                preserved.GetFirstChild<C.AxisId>()?.Remove();
                preserved.PrependChild((C.AxisId)generatedId.CloneNode(true));
            }
            if (generatedCrossing != null) {
                C.CrossingAxis? crossing = preserved.GetFirstChild<C.CrossingAxis>();
                if (crossing != null) crossing.Val = generatedCrossing.Val;
            }
            replacement.ReplaceChild(preserved, generated);
        }
    }
}
