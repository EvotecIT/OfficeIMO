using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a chart on a worksheet.
    /// </summary>
    public sealed partial class ExcelChart {
        private bool ApplySeriesByIndex(int seriesIndex, Action<OpenXmlCompositeElement> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesByIndex(plotArea.Elements<C.BarChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.Bar3DChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.LineChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.Line3DChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.AreaChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.Area3DChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.PieChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.Pie3DChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.OfPieChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.DoughnutChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.ScatterChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.BubbleChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.RadarChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.StockChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.Surface3DChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesByIndex(plotArea.Elements<C.SurfaceChart>(), seriesIndex, apply)) return true;

            return false;
        }

        private bool ApplySeriesByChartIndex(int seriesIndex, Action<OpenXmlCompositeElement> apply) {
            OpenXmlCompositeElement? series = FindSeriesByChartIndex(seriesIndex);
            if (series == null) return false;
            apply(series);
            return true;
        }

        private OpenXmlCompositeElement? FindSeriesByChartIndex(int seriesIndex) {
            C.PlotArea? plotArea = GetChart().GetFirstChild<C.PlotArea>();
            return plotArea?.Descendants()
                .OfType<OpenXmlCompositeElement>()
                .FirstOrDefault(element => IsSeriesElement(element) && GetSeriesIndex(element) == seriesIndex);
        }

        private bool ApplySeriesMarkerByChartIndex(int seriesIndex, Action<C.Marker> apply) {
            OpenXmlCompositeElement? series = FindSeriesByChartIndex(seriesIndex);
            if (series == null || !IsMarkerCapableSeriesElement(series)) return false;

            C.Marker marker = series.GetFirstChild<C.Marker>() ?? new C.Marker();
            apply(marker);
            OpenXmlElement? insertBefore = series.GetFirstChild<C.DataPoint>();
            insertBefore ??= series.GetFirstChild<C.DataLabels>();
            insertBefore ??= series.GetFirstChild<C.Trendline>();
            insertBefore ??= series.GetFirstChild<C.ErrorBars>();
            insertBefore ??= series.GetFirstChild<C.CategoryAxisData>();
            insertBefore ??= series.GetFirstChild<C.Values>();
            insertBefore ??= series.GetFirstChild<C.XValues>();
            insertBefore ??= series.GetFirstChild<C.YValues>();
            insertBefore ??= series.GetFirstChild<C.BubbleSize>();
            insertBefore ??= series.GetFirstChild<C.Smooth>();
            insertBefore ??= series.GetFirstChild<C.ExtensionList>();
            EnsureSeriesChildPosition(series, marker, insertBefore);
            return true;
        }

        private bool ApplySeriesByName(string seriesName, bool ignoreCase, Action<OpenXmlCompositeElement> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesByName(plotArea.Elements<C.BarChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.Bar3DChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.LineChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.Line3DChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.AreaChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.Area3DChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.PieChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.Pie3DChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.OfPieChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.DoughnutChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.ScatterChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.BubbleChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.RadarChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.StockChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.Surface3DChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesByName(plotArea.Elements<C.SurfaceChart>(), seriesName, ignoreCase, apply)) return true;

            return false;
        }

        private bool ApplySeriesMarkerByIndex(int seriesIndex, Action<C.Marker> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesMarkerByIndex(plotArea.Elements<C.LineChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesMarkerByIndex(plotArea.Elements<C.Line3DChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesMarkerByIndex(plotArea.Elements<C.ScatterChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesMarkerByIndex(plotArea.Elements<C.RadarChart>(), seriesIndex, apply)) return true;
            if (ApplySeriesMarkerByIndex(plotArea.Elements<C.StockChart>(), seriesIndex, apply)) return true;

            return false;
        }

        private bool ApplySeriesMarkerByName(string seriesName, bool ignoreCase, Action<C.Marker> apply) {
            C.Chart chart = GetChart();
            C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                return false;
            }

            if (ApplySeriesMarkerByName(plotArea.Elements<C.LineChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesMarkerByName(plotArea.Elements<C.Line3DChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesMarkerByName(plotArea.Elements<C.ScatterChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesMarkerByName(plotArea.Elements<C.RadarChart>(), seriesName, ignoreCase, apply)) return true;
            if (ApplySeriesMarkerByName(plotArea.Elements<C.StockChart>(), seriesName, ignoreCase, apply)) return true;

            return false;
        }

        private static bool ApplySeriesByIndex<TChart>(IEnumerable<TChart> charts, int seriesIndex,
            Action<OpenXmlCompositeElement> apply) where TChart : OpenXmlCompositeElement {
            foreach (TChart chart in charts) {
                List<OpenXmlCompositeElement> series = chart.ChildElements
                    .OfType<OpenXmlCompositeElement>()
                    .Where(IsSeriesElement)
                    .ToList();

                OpenXmlCompositeElement? match = series.FirstOrDefault(s => GetSeriesIndex(s) == seriesIndex);
                if (match != null) {
                    apply(match);
                    return true;
                }

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

        private static bool ApplySeriesMarkerByIndex<TChart>(IEnumerable<TChart> charts, int seriesIndex, Action<C.Marker> apply) where TChart : OpenXmlCompositeElement {
            foreach (TChart chart in charts) {
                List<OpenXmlCompositeElement> series = chart.ChildElements
                    .OfType<OpenXmlCompositeElement>()
                    .Where(IsSeriesElement)
                    .ToList();

                OpenXmlCompositeElement? seriesElement = series.FirstOrDefault(s => GetSeriesIndex(s) == seriesIndex);
                if (seriesElement == null) {
                    if (seriesIndex < 0 || seriesIndex >= series.Count) {
                        continue;
                    }
                    seriesElement = series[seriesIndex];
                }

                C.Marker marker = seriesElement.GetFirstChild<C.Marker>() ?? new C.Marker();
                apply(marker);
                OpenXmlElement? insertBefore = seriesElement.GetFirstChild<C.DataLabels>();
                insertBefore ??= seriesElement.GetFirstChild<C.Trendline>();
                insertBefore ??= seriesElement.GetFirstChild<C.ErrorBars>();
                insertBefore ??= seriesElement.GetFirstChild<C.CategoryAxisData>();
                insertBefore ??= seriesElement.GetFirstChild<C.Values>();
                insertBefore ??= seriesElement.GetFirstChild<C.XValues>();
                insertBefore ??= seriesElement.GetFirstChild<C.YValues>();
                insertBefore ??= seriesElement.GetFirstChild<C.BubbleSize>();
                insertBefore ??= seriesElement.GetFirstChild<C.Smooth>();
                insertBefore ??= seriesElement.GetFirstChild<C.ExtensionList>();
                EnsureSeriesChildPosition(seriesElement, marker, insertBefore);
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
                        OpenXmlElement? insertBefore = series.GetFirstChild<C.DataLabels>();
                        insertBefore ??= series.GetFirstChild<C.Trendline>();
                        insertBefore ??= series.GetFirstChild<C.ErrorBars>();
                        insertBefore ??= series.GetFirstChild<C.CategoryAxisData>();
                        insertBefore ??= series.GetFirstChild<C.Values>();
                        insertBefore ??= series.GetFirstChild<C.XValues>();
                        insertBefore ??= series.GetFirstChild<C.YValues>();
                        insertBefore ??= series.GetFirstChild<C.BubbleSize>();
                        insertBefore ??= series.GetFirstChild<C.Smooth>();
                        insertBefore ??= series.GetFirstChild<C.ExtensionList>();
                        EnsureSeriesChildPosition(series, marker, insertBefore);
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
                   element is C.ScatterChartSeries ||
                   element is C.BubbleChartSeries ||
                   element is C.RadarChartSeries ||
                   element is C.SurfaceChartSeries;
        }

        private static bool IsMarkerCapableSeriesElement(OpenXmlCompositeElement element) =>
            element is C.LineChartSeries ||
            element is C.ScatterChartSeries ||
            element is C.RadarChartSeries;

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
            T? existing = parent.GetFirstChild<T>();
            if (existing != null) {
                parent.ReplaceChild(child, existing);
            } else {
                parent.Append(child);
            }
        }

        private static string NormalizeHexColor(string hex) {
            hex = hex.Trim();
            if (hex.StartsWith("#", StringComparison.Ordinal)) {
                hex = hex.Substring(1);
            }
            if (hex.Length == 6) return hex.ToUpperInvariant();
            if (hex.Length == 8) return hex.Substring(2).ToUpperInvariant();
            return hex.ToUpperInvariant();
        }
    }
}
