using DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word {
    public partial class WordChart {
        public void AddCategories(List<string> categories) {
            Categories = categories;
        }
        public WordChart AddPie<T>(string category, T value) {
            // if value is a list we need to throw as not supported
            if (!(value is int || value is double || value is float)) {
                throw new NotSupportedException("Value must be of type int, double, or float");
            }
            EnsureChartExistsPie();
            AddSingleCategory(category);
            AddSingleValue(value);
            return this;
        }

        public WordChart AddPie3D<T>(string category, T value) {
            if (!(value is int || value is double || value is float)) {
                throw new NotSupportedException("Value must be of type int, double, or float");
            }
            EnsureChartExistsPie3D();
            AddSingleCategory(category);
            AddSingleValue(value);
            return this;
        }

        public void AddChartLine<T>(string name, int[] values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsLine();
            if (_chart != null) {
                var lineChart = _chart.PlotArea.GetFirstChild<LineChart>();
                if (lineChart != null) {
                    LineChartSeries lineChartSeries = AddLineChartSeries(this._index, name, color, this.Categories, values.ToList());
                    InsertSeries(lineChart, lineChartSeries);
                }
            }
        }

        /// <summary>
        /// Add a line to a chart. Multiple lines can be added.
        /// You cannot mix lines with pies or bars.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="values"></param>
        /// <param name="color"></param>
        public void AddLine<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsLine();
            var lineChart = _chart.PlotArea.GetFirstChild<LineChart>();
            if (lineChart != null) {
                LineChartSeries lineChartSeries = AddLineChartSeries(this._index, name, color, this.Categories, values);
                InsertSeries(lineChart, lineChartSeries);
            }

        }

        public void AddChartAxisX(List<string> categories) {
            Categories = categories;
        }

        public void AddBar(string name, int values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar();
            var barChart = _chart.PlotArea.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, new List<int>() { values });
                InsertSeries(barChart, barChartSeries);
            }
        }

        public void AddBar<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar();
            var barChart = _chart.PlotArea.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, values);
                InsertSeries(barChart, barChartSeries);
            }
        }

        public void AddBar(string name, int[] values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar();
            var barChart = _chart.PlotArea.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, values.ToList());
                InsertSeries(barChart, barChartSeries);
            }
        }

        public void AddArea<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsArea();
            if (_chart != null) {
                var barChart = _chart.PlotArea.GetFirstChild<AreaChart>();
                if (barChart != null) {
                    AreaChartSeries areaChartSeries = AddAreaChartSeries(this._index, name, color, this.Categories, values);
                    InsertSeries(barChart, areaChartSeries);
                }
            }
        }

        public void AddArea<T>(string name, int[] values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsArea();
            if (_chart != null) {
                var barChart = _chart.PlotArea.GetFirstChild<AreaChart>();
                if (barChart != null) {
                    AreaChartSeries areaChartSeries = AddAreaChartSeries(this._index, name, color, this.Categories, values.ToList());
                    InsertSeries(barChart, areaChartSeries);
                }
            }
        }

        public void AddScatter(string name, List<double> xValues, List<double> yValues, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsScatter();
            if (_chart != null) {
                var scatterChart = _chart.PlotArea.GetFirstChild<ScatterChart>();
                if (scatterChart != null) {
                    var series = AddScatterChartSeries(this._index, name, color, xValues, yValues);
                    InsertSeries(scatterChart, series);
                }
            }
        }

        public void AddRadar<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsRadar();
            if (_chart != null) {
                var radarChart = _chart.PlotArea.GetFirstChild<RadarChart>();
                if (radarChart != null) {
                    var series = AddRadarChartSeries(this._index, name, color, this.Categories, values);
                    InsertSeries(radarChart, series);
                }
            }
        }
        public void AddBar3D<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar3D();
            if (_chart != null) {
                var chart3d = _chart.PlotArea.GetFirstChild<Bar3DChart>();
                if (chart3d != null) {
                    var series = AddBar3DChartSeries(this._index, name, color, this.Categories, values);

                    // For Bar3DChart, we need special handling to maintain correct element order:
                    // barDir, grouping, varyColors, ser, dLbls, gapWidth, gapDepth, shape, axId, extLst
                    var axis = chart3d.Elements<AxisId>().FirstOrDefault();
                    if (axis != null) {
                        chart3d.InsertBefore(series, axis);

                        // Ensure gapWidth is present and in correct position (after all ser elements, before axId)
                        var gapWidth = chart3d.GetFirstChild<GapWidth>();
                        if (gapWidth == null) {
                            gapWidth = new GapWidth() { Val = (UInt16Value)150U };
                            chart3d.InsertBefore(gapWidth, axis);
                        }
                    } else {
                        chart3d.Append(series);
                    }
                }
            }
        }

        public void AddLine3D<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsLine3D();
            if (_chart != null) {
                var line3d = _chart.PlotArea.GetFirstChild<Line3DChart>();
                if (line3d != null) {
                    var series = AddLine3DChartSeries(this._index, name, color, this.Categories, values);
                    InsertSeries(line3d, series);
                }
            }
        }

        public void AddLegend(LegendPositionValues legendPosition) {
            if (_chart != null) {
                Legend legend = new Legend();
                LegendPosition postion = new LegendPosition() { Val = legendPosition };
                Overlay overlay = new Overlay() { Val = false };
                legend.Append(postion);
                legend.Append(overlay);

                // Insert legend in correct position according to OpenXML schema
                // Legend should come after PlotArea but before other elements like PlotVisibleOnly
                var plotVisibleOnly = _chart.GetFirstChild<PlotVisibleOnly>();
                if (plotVisibleOnly != null) {
                    _chart.InsertBefore(legend, plotVisibleOnly);
                } else {
                    // If no PlotVisibleOnly, just append at the end
                    _chart.Append(legend);
                }
            }
        }
    }
}
