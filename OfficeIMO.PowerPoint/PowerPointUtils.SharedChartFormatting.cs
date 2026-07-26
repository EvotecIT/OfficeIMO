using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
        private static void PreserveSharedChartFormatting(C.PlotArea source, C.PlotArea replacement) {
            PreserveSharedChartLayers(source, replacement);
            PreserveSharedAxes(source, replacement);
        }

        private static void PreserveSharedChartLayers(C.PlotArea source, C.PlotArea replacement) {
            List<OpenXmlCompositeElement> sourceLayers = source.ChildElements
                .OfType<OpenXmlCompositeElement>().Where(IsSharedChartLayer).ToList();
            var usedLayers = new HashSet<OpenXmlCompositeElement>();
            foreach (OpenXmlCompositeElement generated in replacement.ChildElements
                         .OfType<OpenXmlCompositeElement>().Where(IsSharedChartLayer).ToList()) {
                OpenXmlCompositeElement? match = sourceLayers.FirstOrDefault(candidate =>
                    !usedLayers.Contains(candidate) &&
                    AreCompatibleSharedChartLayers(candidate, generated, source, replacement));
                if (match == null) continue;

                usedLayers.Add(match);
                var preserved = (OpenXmlCompositeElement)match.CloneNode(true);
                ReplaceSharedSeriesData(preserved, generated);
                ReplaceSharedAxisReferences(preserved, generated);
                replacement.ReplaceChild(preserved, generated);
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
            } else if (source is C.RadarChart sourceRadar && replacement is C.RadarChart replacementRadar &&
                       sourceRadar.RadarStyle?.Val?.Value != replacementRadar.RadarStyle?.Val?.Value) {
                return false;
            }

            return IsSecondarySharedChartLayer(source, sourcePlotArea) ==
                   IsSecondarySharedChartLayer(replacement, replacementPlotArea);
        }

        private static bool IsSecondarySharedChartLayer(OpenXmlCompositeElement chartLayer,
            C.PlotArea plotArea) {
            C.AxisId? categoryReference = chartLayer.Elements<C.AxisId>().FirstOrDefault();
            if (categoryReference?.Val == null) return false;
            OpenXmlCompositeElement? categoryAxis = plotArea.ChildElements
                .OfType<OpenXmlCompositeElement>()
                .FirstOrDefault(axis => IsSharedCategoryAxis(axis) &&
                    axis.GetFirstChild<C.AxisId>()?.Val?.Value == categoryReference.Val.Value);
            return categoryAxis?.GetFirstChild<C.Delete>()?.Val?.Value == true;
        }

        private static void ReplaceSharedSeriesData(OpenXmlCompositeElement preserved,
            OpenXmlCompositeElement generated) {
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
                if (sourceSeries != null) usedSeries.Add(sourceSeries);
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

        private static void PreserveSharedAxes(C.PlotArea source, C.PlotArea replacement) {
            if (UsesHorizontalSharedAxes(source) != UsesHorizontalSharedAxes(replacement)) return;
            PreserveSharedCategoryAxes(source, replacement);
            if (!PreserveSharedBubbleAxes(source, replacement)) {
                PreserveSharedAxes<C.ValueAxis>(source, replacement);
            }
        }

        private static bool PreserveSharedBubbleAxes(
            C.PlotArea source, C.PlotArea replacement) {
            C.BubbleChart? sourceChart = source.GetFirstChild<C.BubbleChart>();
            C.BubbleChart? replacementChart =
                replacement.GetFirstChild<C.BubbleChart>();
            if (sourceChart == null || replacementChart == null) return false;

            List<C.AxisId> sourceReferences =
                sourceChart.Elements<C.AxisId>().ToList();
            List<C.AxisId> replacementReferences =
                replacementChart.Elements<C.AxisId>().ToList();
            if (sourceReferences.Count != replacementReferences.Count ||
                sourceReferences.Count == 0) {
                return false;
            }

            var axisPairs =
                new List<(C.ValueAxis Source, C.ValueAxis Replacement)>();
            for (int index = 0; index < sourceReferences.Count; index++) {
                uint? sourceId = sourceReferences[index].Val?.Value;
                uint? replacementId = replacementReferences[index].Val?.Value;
                C.ValueAxis? sourceAxis = source.Elements<C.ValueAxis>()
                    .FirstOrDefault(axis =>
                        axis.AxisId?.Val?.Value == sourceId);
                C.ValueAxis? replacementAxis =
                    replacement.Elements<C.ValueAxis>().FirstOrDefault(axis =>
                        axis.AxisId?.Val?.Value == replacementId);
                if (sourceAxis == null || replacementAxis == null) return false;
                axisPairs.Add((sourceAxis, replacementAxis));
            }
            foreach ((C.ValueAxis sourceAxis,
                      C.ValueAxis replacementAxis) in axisPairs) {
                ReplaceSharedAxis(replacement, sourceAxis, replacementAxis);
            }
            return true;
        }

        private static bool IsSharedCategoryAxis(OpenXmlCompositeElement axis) =>
            axis is C.CategoryAxis || axis is C.DateAxis;

        private static void PreserveSharedCategoryAxes(C.PlotArea source, C.PlotArea replacement) {
            List<OpenXmlCompositeElement> sourceAxes = source.ChildElements
                .OfType<OpenXmlCompositeElement>().Where(IsSharedCategoryAxis).ToList();
            List<OpenXmlCompositeElement> replacementAxes = replacement.ChildElements
                .OfType<OpenXmlCompositeElement>().Where(IsSharedCategoryAxis).ToList();
            int count = Math.Min(sourceAxes.Count, replacementAxes.Count);
            for (int index = 0; index < count; index++) {
                ReplaceSharedAxis(replacement, sourceAxes[index], replacementAxes[index]);
            }
        }

        private static bool UsesHorizontalSharedAxes(C.PlotArea plotArea) =>
            plotArea.Elements<C.BarChart>().Any(chart =>
                chart.BarDirection?.Val?.Value == C.BarDirectionValues.Bar);

        private static void PreserveSharedAxes<TAxis>(C.PlotArea source, C.PlotArea replacement)
            where TAxis : OpenXmlCompositeElement {
            List<TAxis> sourceAxes = source.Elements<TAxis>().ToList();
            List<TAxis> replacementAxes = replacement.Elements<TAxis>().ToList();
            int count = Math.Min(sourceAxes.Count, replacementAxes.Count);
            for (int index = 0; index < count; index++) {
                ReplaceSharedAxis(replacement, sourceAxes[index], replacementAxes[index]);
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
