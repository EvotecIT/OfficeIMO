using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using ChartIndex = DocumentFormat.OpenXml.Drawing.Charts.Index;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        internal static void UpdateChartData(ChartPart chartPart, ExcelChartData data, ExcelChartDataRange range) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            ChartSpace? chartSpace = chartPart.ChartSpace;
            Chart? chart = chartSpace?.GetFirstChild<Chart>();
            PlotArea? plotArea = chart?.GetFirstChild<PlotArea>();
            if (plotArea == null) {
                throw new InvalidOperationException("Chart plot area not found.");
            }

            int chartElementCount =
                plotArea.Elements<BarChart>().Count()
                + plotArea.Elements<Bar3DChart>().Count()
                + plotArea.Elements<LineChart>().Count()
                + plotArea.Elements<Line3DChart>().Count()
                + plotArea.Elements<AreaChart>().Count()
                + plotArea.Elements<Area3DChart>().Count()
                + plotArea.Elements<PieChart>().Count()
                + plotArea.Elements<Pie3DChart>().Count()
                + plotArea.Elements<OfPieChart>().Count()
                + plotArea.Elements<DoughnutChart>().Count()
                + plotArea.Elements<ScatterChart>().Count()
                + plotArea.Elements<BubbleChart>().Count()
                + plotArea.Elements<RadarChart>().Count()
                + plotArea.Elements<StockChart>().Count()
                + plotArea.Elements<Surface3DChart>().Count()
                + plotArea.Elements<SurfaceChart>().Count();

            ExcelChartType defaultType = InferChartType(plotArea);
            List<SeriesDescriptor> descriptors = BuildSeriesDescriptors(range, data, defaultType, useSeriesOverrides: chartElementCount > 1);
            ValidateSingleSeriesPieVariants(descriptors);

            if (plotArea.GetFirstChild<ScatterChart>() is ScatterChart scatterChart) {
                if (chartElementCount > 1) {
                    UpdateComboChartData(plotArea, data, range, descriptors);
                } else {
                    UpdateScatterChartSeries(scatterChart, data, range, descriptors);
                }
                return;
            }

            if (plotArea.GetFirstChild<BubbleChart>() != null) {
                throw new NotSupportedException("Updating bubble charts is not supported. Use range-based charts for bubble data.");
            }
            if (plotArea.GetFirstChild<StockChart>() is StockChart stockChart) {
                if (chartElementCount > 1) {
                    throw new NotSupportedException("Stock charts cannot be updated as combination charts.");
                }
                UpdateStockChartSeries(stockChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<Surface3DChart>() is Surface3DChart surface3DChart) {
                if (chartElementCount > 1) {
                    throw new NotSupportedException("Surface charts cannot be updated as combination charts.");
                }
                UpdateSurfaceChartSeries(surface3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<SurfaceChart>() is SurfaceChart surfaceChart) {
                if (chartElementCount > 1) {
                    throw new NotSupportedException("Surface charts cannot be updated as combination charts.");
                }
                UpdateSurfaceChartSeries(surfaceChart, data, range, descriptors);
                return;
            }

            if (chartElementCount > 1) {
                UpdateComboChartData(plotArea, data, range, descriptors);
                return;
            }

            if (plotArea.GetFirstChild<BarChart>() is BarChart barChart) {
                UpdateBarChartSeries(barChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<Bar3DChart>() is Bar3DChart bar3DChart) {
                UpdateBar3DChartSeries(bar3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<LineChart>() is LineChart lineChart) {
                UpdateLineChartSeries(lineChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<Line3DChart>() is Line3DChart line3DChart) {
                UpdateLine3DChartSeries(line3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<AreaChart>() is AreaChart areaChart) {
                UpdateAreaChartSeries(areaChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<Area3DChart>() is Area3DChart area3DChart) {
                UpdateArea3DChartSeries(area3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<RadarChart>() is RadarChart radarChart) {
                UpdateRadarChartSeries(radarChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<PieChart>() is PieChart pieChart) {
                UpdatePieChartSeries(pieChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<Pie3DChart>() is Pie3DChart pie3DChart) {
                UpdatePie3DChartSeries(pie3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<OfPieChart>() is OfPieChart ofPieChart) {
                UpdateOfPieChartSeries(ofPieChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<DoughnutChart>() is DoughnutChart doughnutChart) {
                UpdateDoughnutChartSeries(doughnutChart, data, range, descriptors);
                return;
            }

            throw new NotSupportedException("Chart type is not supported for data updates.");
        }
    }
}
