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

        public void AddChartLine<T>(string name, int[] values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsLine();
            if (_chart != null) {
                var lineChart = _chart.PlotArea.GetFirstChild<LineChart>();
                if (lineChart != null) {
                    LineChartSeries lineChartSeries = AddLineChartSeries(this._index, name, color, this.Categories, values.ToList());
                    lineChart.Append(lineChartSeries);
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
                lineChart.Append(lineChartSeries);
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
                barChart.Append(barChartSeries);
            }
        }

        public void AddBar<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar();
            var barChart = _chart.PlotArea.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, values);
                barChart.Append(barChartSeries);
            }
        }

        public void AddBar(string name, int[] values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar();
            var barChart = _chart.PlotArea.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, values.ToList());
                barChart.Append(barChartSeries);
            }
        }

        public void AddArea<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsArea();
            if (_chart != null) {
                var barChart = _chart.PlotArea.GetFirstChild<AreaChart>();
                if (barChart != null) {
                    AreaChartSeries areaChartSeries = AddAreaChartSeries(this._index, name, color, this.Categories, values);
                    barChart.Append(areaChartSeries);
                }
            }
        }

        public void AddArea<T>(string name, int[] values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsArea();
            if (_chart != null) {
                var barChart = _chart.PlotArea.GetFirstChild<AreaChart>();
                if (barChart != null) {
                    AreaChartSeries areaChartSeries = AddAreaChartSeries(this._index, name, color, this.Categories, values.ToList());
                    barChart.Append(areaChartSeries);
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
                _chart.Append(legend);
            }
        }
    }
}
