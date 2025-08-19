using DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word {
    /// <summary>
    /// Partial class containing public methods for building chart content.
    /// </summary>
    public partial class WordChart {
        /// <summary>
        /// Sets the category labels used by subsequent chart series.
        /// </summary>
        /// <param name="categories">List of category names.</param>
        public void AddCategories(List<string> categories) {
            Categories = categories;
        }
        /// <summary>
        /// Adds a single value to a pie chart.
        /// </summary>
        /// <typeparam name="T">Numeric type representing the value.</typeparam>
        /// <param name="category">Category label.</param>
        /// <param name="value">Data value for the slice.</param>
        /// <returns>The current <see cref="WordChart"/> instance.</returns>
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

        /// <summary>
        /// Adds a single value to a 3D pie chart.
        /// </summary>
        /// <typeparam name="T">Numeric type representing the value.</typeparam>
        /// <param name="category">Category label.</param>
        /// <param name="value">Data value for the slice.</param>
        /// <returns>The current <see cref="WordChart"/> instance.</returns>
        public WordChart AddPie3D<T>(string category, T value) {
            if (!(value is int || value is double || value is float)) {
                throw new NotSupportedException("Value must be of type int, double, or float");
            }
            EnsureChartExistsPie3D();
            AddSingleCategory(category);
            AddSingleValue(value);
            return this;
        }

        /// <summary>
        /// Adds a line series to the chart from an array of integer values.
        /// </summary>
        /// <typeparam name="T">Unused generic parameter.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Values for the series.</param>
        /// <param name="color">Line color.</param>
        public void AddChartLine<T>(string name, int[] values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsLine();
            var lineChart = _chart?.PlotArea?.GetFirstChild<LineChart>();
            if (lineChart != null) {
                LineChartSeries lineChartSeries = AddLineChartSeries(this._index, name, color, this.Categories, values.ToList());
                InsertSeries(lineChart, lineChartSeries);
            }
        }

        /// <summary>
        /// Add a line to a chart. Multiple lines can be added.
        /// You cannot mix lines with pie charts. Lines can be combined with bar
        /// series to create combo charts, but make sure to call
        /// <see cref="AddChartAxisX"/> before adding either series.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="values"></param>
        /// <param name="color"></param>
        public void AddLine<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsLine();
            var lineChart = _chart?.PlotArea?.GetFirstChild<LineChart>();
            if (lineChart != null) {
                LineChartSeries lineChartSeries = AddLineChartSeries(this._index, name, color, this.Categories, values);
                InsertSeries(lineChart, lineChartSeries);
            }

        }

        /// <summary>
        /// Sets the category labels for the X axis. This should be called
        /// before adding bar or line series when creating a combo chart so that
        /// both chart types share the same axis.
        /// </summary>
        public void AddChartAxisX(List<string> categories) {
            Categories = categories;
        }

        /// <summary>
        /// Adds a bar series to the chart. When mixing bar and line series be
        /// sure to call <see cref="AddChartAxisX"/> first so the categories are
        /// shared across both chart types.
        /// </summary>
        public void AddBar(string name, int values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar();
            var barChart = _chart?.PlotArea?.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, new List<int>() { values });
                InsertSeries(barChart, barChartSeries);
            }
        }

        /// <summary>
        /// Adds a bar series with multiple values. When used in a combo chart
        /// with line series the categories must be set via
        /// <see cref="AddChartAxisX"/> before calling this method.
        /// </summary>
        public void AddBar<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar();
            var barChart = _chart?.PlotArea?.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, values);
                InsertSeries(barChart, barChartSeries);
            }
        }

        /// <summary>
        /// Adds a bar series from an array of values. For combo charts ensure
        /// <see cref="AddChartAxisX"/> has been called first.
        /// </summary>
        public void AddBar(string name, int[] values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar();
            var barChart = _chart.PlotArea.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, values.ToList());
                InsertSeries(barChart, barChartSeries);
            }
        }

        /// <summary>
        /// Adds an area chart series using the provided values.
        /// </summary>
        /// <typeparam name="T">Numeric type of the data values.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Values for the series.</param>
        /// <param name="color">Fill color for the series.</param>
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

        /// <summary>
        /// Adds an area chart series using an array of integer values.
        /// </summary>
        /// <typeparam name="T">Unused generic parameter.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Data values.</param>
        /// <param name="color">Fill color for the series.</param>
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

        /// <summary>
        /// Adds a scatter chart series with separate X and Y values.
        /// </summary>
        /// <param name="name">Series name.</param>
        /// <param name="xValues">Values plotted on the X axis.</param>
        /// <param name="yValues">Values plotted on the Y axis.</param>
        /// <param name="color">Color of the series.</param>
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

        /// <summary>
        /// Adds a radar chart series.
        /// </summary>
        /// <typeparam name="T">Numeric type of the values.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Values for the series.</param>
        /// <param name="color">Color of the series.</param>
        public void AddRadar<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsRadar();
            var radarChart = _chart?.PlotArea?.GetFirstChild<RadarChart>();
            if (radarChart != null) {
                var series = AddRadarChartSeries(this._index, name, color, this.Categories, values);
                InsertSeries(radarChart, series);
            }
        }
        /// <summary>
        /// Adds a 3D bar chart series.
        /// </summary>
        /// <typeparam name="T">Numeric type of the data values.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Series data values.</param>
        /// <param name="color">Color of the bars.</param>
        public void AddBar3D<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar3D();
            var chart3d = _chart?.PlotArea?.GetFirstChild<Bar3DChart>();
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

        /// <summary>
        /// Adds a 3D line chart series to the chart.
        /// </summary>
        /// <param name="name">The name of the series</param>
        /// <param name="values">The data values for the series</param>
        /// <param name="color">The color of the series</param>
        /// <typeparam name="T">The type of data values</typeparam>
        /// <remarks>
        /// KNOWN ISSUE: Line3DChart currently fails OpenXML schema validation with the error:
        /// "The element has unexpected child element 'ser'". This appears to be a discrepancy
        /// between Microsoft's documentation (which shows Line3DChart supports LineChartSeries)
        /// and the actual OpenXML schema validator implementation. This issue persists regardless
        /// of element ordering and may indicate that Line3DChart does not actually support data
        /// series in the current OpenXML specification.
        /// </remarks>
        public void AddLine3D<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsLine3D();
            var line3d = _chart?.PlotArea?.GetFirstChild<Line3DChart>();
            if (line3d != null) {
                var series = AddLine3DChartSeries(this._index, name, color, this.Categories, values);
                // Insert series in the correct schema position (after varyColors, before dLbls)
                InsertSeries(line3d, series);
            }
        }

        /// <summary>
        /// Adds a 3D area chart series.
        /// </summary>
        /// <typeparam name="T">Numeric type of the data values.</typeparam>
        /// <param name="name">Series name.</param>
        /// <param name="values">Series data values.</param>
        /// <param name="color">Series color.</param>
        public void AddArea3D<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsArea3D();
            var area3d = _chart?.PlotArea?.GetFirstChild<Area3DChart>();
            if (area3d != null) {
                var series = AddArea3DChartSeries(this._index, name, color, this.Categories, values);
                InsertSeries(area3d, series);
            }
        }

        /// <summary>
        /// Adds a legend to the chart at the specified position.
        /// </summary>
        /// <param name="legendPosition">Desired legend position.</param>
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

        /// <summary>
        /// Sets the title displayed on the X axis.
        /// </summary>
        /// <param name="title">Axis title text.</param>
        /// <returns>The current <see cref="WordChart"/> instance.</returns>
        public WordChart SetXAxisTitle(string title) {
            _xAxisTitle = title;
            UpdateAxisTitles();
            return this;
        }

        /// <summary>
        /// Sets the title displayed on the Y axis.
        /// </summary>
        /// <param name="title">Axis title text.</param>
        /// <returns>The current <see cref="WordChart"/> instance.</returns>
        public WordChart SetYAxisTitle(string title) {
            _yAxisTitle = title;
            UpdateAxisTitles();
            return this;
        }

        /// <summary>
        /// Defines font formatting for axis titles.
        /// </summary>
        /// <param name="fontName">Font family name.</param>
        /// <param name="fontSize">Font size in points.</param>
        /// <param name="color">Text color.</param>
        /// <returns>The current <see cref="WordChart"/> instance.</returns>
        public WordChart SetAxisTitleFormat(string fontName, int fontSize, SixLabors.ImageSharp.Color color) {
            _axisTitleFontName = fontName;
            _axisTitleFontSize = fontSize;
            _axisTitleColor = color;
            UpdateAxisTitles();
            return this;
        }
    }
}
