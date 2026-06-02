using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
        /// <summary>
        ///     Represents a chart on a slide.
        /// </summary>
        public partial class PowerPointChart : PowerPointShape {
        private bool ApplySeriesByIndex(int seriesIndex, Action<OpenXmlCompositeElement> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesByIndex(plotArea.Elements<C.BarChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.LineChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.AreaChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.PieChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.DoughnutChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.ScatterChart>(), seriesIndex, apply)) return true;

            return false;
        }

        private bool ApplySeriesByName(string seriesName, bool ignoreCase, Action<OpenXmlCompositeElement> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesByName(plotArea.Elements<C.BarChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.LineChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.AreaChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.PieChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.DoughnutChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.ScatterChart>(), seriesName, ignoreCase, apply)) return true;

            return false;
        }

        private bool ApplySeriesMarkerByIndex(int seriesIndex, Action<C.Marker> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesMarkerByIndex(plotArea.Elements<C.LineChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesMarkerByIndex(plotArea.Elements<C.ScatterChart>(), seriesIndex, apply)) return true;

            return false;
        }

        private bool ApplySeriesMarkerByName(string seriesName, bool ignoreCase, Action<C.Marker> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesMarkerByName(plotArea.Elements<C.LineChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesMarkerByName(plotArea.Elements<C.ScatterChart>(), seriesName, ignoreCase, apply)) return true;

            return false;
        }

        private static bool ApplySeriesMarkerByIndex<TChart>(IEnumerable<TChart> charts, int seriesIndex, Action<C.Marker> apply) where TChart : OpenXmlCompositeElement {
            foreach (TChart chart in charts) {
                List<OpenXmlCompositeElement> series = chart.ChildElements
                    .OfType<OpenXmlCompositeElement>()
                    .Where(IsSeriesElement)
                    .OrderBy(GetSeriesIndex)
                    .ToList();

                if (seriesIndex < 0 || seriesIndex >= series.Count) {
                    continue;
                }

                OpenXmlCompositeElement seriesElement = series[seriesIndex];
                C.Marker marker = seriesElement.GetFirstChild<C.Marker>() ?? new C.Marker();
                apply(marker);
                if (marker.Parent == null) {
                    InsertSeriesMarker(seriesElement, marker);
                }
                return true;
            }

            return false;
        }

        private static bool ApplySeriesMarkerByName<TChart>(IEnumerable<TChart> charts, string seriesName, bool ignoreCase, Action<C.Marker> apply) where TChart : OpenXmlCompositeElement {
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (TChart chart in charts) {
                foreach (OpenXmlCompositeElement series in chart.ChildElements.OfType<OpenXmlCompositeElement>().Where(IsSeriesElement)) {
                    string? name = GetSeriesName(series);
                    if (name != null && string.Equals(name, seriesName, comparison)) {
                        C.Marker marker = series.GetFirstChild<C.Marker>() ?? new C.Marker();
                        apply(marker);
                        if (marker.Parent == null) {
                            InsertSeriesMarker(series, marker);
                        }
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ApplySeriesByIndex<TChart>(IEnumerable<TChart> charts, int seriesIndex,
            Action<OpenXmlCompositeElement> apply) where TChart : OpenXmlCompositeElement {
            foreach (TChart chart in charts) {
                List<OpenXmlCompositeElement> series = chart.ChildElements
                    .OfType<OpenXmlCompositeElement>()
                    .Where(IsSeriesElement)
                    .OrderBy(GetSeriesIndex)
                    .ToList();

                if (seriesIndex < 0 || seriesIndex >= series.Count) {
                    continue;
                }

                apply(series[seriesIndex]);
                return true;
            }

            return false;
        }

        private static bool ApplySeriesByName<TChart>(IEnumerable<TChart> charts, string seriesName, bool ignoreCase,
            Action<OpenXmlCompositeElement> apply) where TChart : OpenXmlCompositeElement {
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (TChart chart in charts) {
                foreach (OpenXmlCompositeElement series in chart.ChildElements.OfType<OpenXmlCompositeElement>().Where(IsSeriesElement)) {
                    string? name = GetSeriesName(series);
                    if (name != null && string.Equals(name, seriesName, comparison)) {
                        apply(series);
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsSeriesElement(OpenXmlCompositeElement element) {
            return element is C.BarChartSeries ||
                   element is C.LineChartSeries ||
                   element is C.AreaChartSeries ||
                   element is C.PieChartSeries ||
                   element is C.ScatterChartSeries;
        }

        private static int GetSeriesIndex(OpenXmlCompositeElement series) {
            return (int)(series.GetFirstChild<C.Index>()?.Val?.Value ?? 0U);
        }

        private static string? GetSeriesName(OpenXmlCompositeElement series) {
            C.SeriesText? seriesText = series.GetFirstChild<C.SeriesText>();
            if (seriesText == null) {
                return null;
            }

            C.StringReference? reference = seriesText.GetFirstChild<C.StringReference>();
            C.StringCache? cache = reference?.GetFirstChild<C.StringCache>();
            string? cachedText = cache?.Elements<C.StringPoint>()
                .FirstOrDefault()?
                .NumericValue?
                .Text;
            if (!string.IsNullOrWhiteSpace(cachedText)) {
                return cachedText;
            }

            C.StringLiteral? literal = seriesText.GetFirstChild<C.StringLiteral>();
            string? literalText = literal?.Elements<C.StringPoint>()
                .FirstOrDefault()?
                .NumericValue?
                .Text;
            if (!string.IsNullOrWhiteSpace(literalText)) {
                return literalText;
            }

            return string.IsNullOrWhiteSpace(seriesText.InnerText) ? null : seriesText.InnerText;
        }

        private static C.ChartText CreateChartText(string title) {
            return new C.ChartText(
                new C.RichText(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties { Language = "en-US" },
                            new A.Text { Text = title })
                    )));
        }

        private static C.Title CreateAxisTitle(string title) {
            return new C.Title(
                new C.ChartText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(
                                new A.RunProperties { Language = "en-US" },
                                new A.Text { Text = title })))
                ),
                new C.Layout(),
                new C.Overlay { Val = false }
            );
        }

        private static void ReplaceChild<T>(OpenXmlCompositeElement parent, T child) where T : OpenXmlElement {
            parent.GetFirstChild<T>()?.Remove();
            parent.Append(child);
        }

        private static void ReplaceAxisChild<T>(OpenXmlCompositeElement axis, T child) where T : OpenXmlElement {
            axis.GetFirstChild<T>()?.Remove();

            OpenXmlElement? insertBefore = axis.GetFirstChild<C.ShapeProperties>();
            insertBefore ??= axis.GetFirstChild<C.TextProperties>();
            insertBefore ??= axis.GetFirstChild<C.CrossingAxis>();
            insertBefore ??= axis.GetFirstChild<C.Crosses>();
            insertBefore ??= axis.GetFirstChild<C.CrossesAt>();
            insertBefore ??= axis.GetFirstChild<C.AutoLabeled>();
            insertBefore ??= axis.GetFirstChild<C.LabelAlignment>();
            insertBefore ??= axis.GetFirstChild<C.LabelOffset>();
            insertBefore ??= axis.GetFirstChild<C.NoMultiLevelLabels>();
            insertBefore ??= axis.GetFirstChild<C.CrossBetween>();
            insertBefore ??= axis.GetFirstChild<C.MajorUnit>();
            insertBefore ??= axis.GetFirstChild<C.MinorUnit>();
            insertBefore ??= axis.GetFirstChild<C.DisplayUnits>();
            insertBefore ??= axis.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                axis.InsertBefore(child, insertBefore);
            } else {
                axis.Append(child);
            }
        }

        private static void InsertAxisGridlines<TGridlines>(OpenXmlCompositeElement axis, TGridlines gridlines)
            where TGridlines : OpenXmlCompositeElement {
            OpenXmlElement? insertBefore = typeof(TGridlines) == typeof(C.MajorGridlines)
                ? axis.GetFirstChild<C.MinorGridlines>()
                : null;
            insertBefore ??= axis.GetFirstChild<C.Title>();
            insertBefore ??= axis.GetFirstChild<C.NumberingFormat>();
            insertBefore ??= axis.GetFirstChild<C.MajorTickMark>();
            insertBefore ??= axis.GetFirstChild<C.MinorTickMark>();
            insertBefore ??= axis.GetFirstChild<C.TickLabelPosition>();
            insertBefore ??= axis.GetFirstChild<C.ShapeProperties>();
            insertBefore ??= axis.GetFirstChild<C.TextProperties>();
            insertBefore ??= axis.GetFirstChild<C.CrossingAxis>();
            insertBefore ??= axis.GetFirstChild<C.Crosses>();
            insertBefore ??= axis.GetFirstChild<C.CrossesAt>();
            insertBefore ??= axis.GetFirstChild<C.AutoLabeled>();
            insertBefore ??= axis.GetFirstChild<C.LabelAlignment>();
            insertBefore ??= axis.GetFirstChild<C.LabelOffset>();
            insertBefore ??= axis.GetFirstChild<C.NoMultiLevelLabels>();
            insertBefore ??= axis.GetFirstChild<C.CrossBetween>();
            insertBefore ??= axis.GetFirstChild<C.MajorUnit>();
            insertBefore ??= axis.GetFirstChild<C.MinorUnit>();
            insertBefore ??= axis.GetFirstChild<C.DisplayUnits>();
            insertBefore ??= axis.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                axis.InsertBefore(gridlines, insertBefore);
            } else {
                axis.Append(gridlines);
            }
        }

        private static void ReplaceValueAxisCrossBetween(C.ValueAxis axis, C.CrossBetween crossBetween) {
            axis.GetFirstChild<C.CrossBetween>()?.Remove();

            OpenXmlElement? insertBefore = axis.GetFirstChild<C.MajorUnit>();
            insertBefore ??= axis.GetFirstChild<C.MinorUnit>();
            insertBefore ??= axis.GetFirstChild<C.DisplayUnits>();
            insertBefore ??= axis.GetFirstChild<C.ExtensionList>();

            if (insertBefore != null) {
                axis.InsertBefore(crossBetween, insertBefore);
            } else {
                axis.Append(crossBetween);
            }
        }
    }
}
