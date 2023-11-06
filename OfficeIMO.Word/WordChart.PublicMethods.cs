using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Drawing.Charts;
using SixLabors.ImageSharp;

namespace OfficeIMO.Word {
    public partial class WordChart {
        public void AddCategories(List<string> categories) {
            Categories = categories;
        }

        public void AddChartPie<T>(string name, int[] values) {
            if (_chart != null) {
                var pieChart = _chart.PlotArea.GetFirstChild<PieChart>();
                if (pieChart != null) {
                    PieChartSeries pieChartSeries = WordPieChart.AddPieChartSeries(this._index, name, this.Categories, values.ToList());
                    pieChart.Append(pieChartSeries);
                }
            }
        }

        public void AddChartPie<T>(string name, List<T> values) {
            if (_chart != null) {
                var pieChart = _chart.PlotArea.GetFirstChild<PieChart>();
                if (pieChart != null) {
                    PieChartSeries pieChartSeries = WordPieChart.AddPieChartSeries(this._index, name, this.Categories, values);
                    pieChart.Append(pieChartSeries);
                }
            }
        }

        public void AddChartLine<T>(string name, int[] values, SixLabors.ImageSharp.Color color) {
            if (_chart != null) {
                var lineChart = _chart.PlotArea.GetFirstChild<LineChart>();
                if (lineChart != null) {
                    LineChartSeries lineChartSeries = WordLineChart.AddLineChartSeries(this._index, name, color, this.Categories, values.ToList());
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
        public void AddChartLine<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            if (_chart != null) {
                var lineChart = _chart.PlotArea.GetFirstChild<LineChart>();
                if (lineChart != null) {
                    LineChartSeries lineChartSeries = WordLineChart.AddLineChartSeries(this._index, name, color, this.Categories, values);
                    lineChart.Append(lineChartSeries);
                }
            }
        }

        public void AddChartAxisX(List<string> categories) {
            Categories = categories;
        }

        public void AddChartBar(string name, int values, SixLabors.ImageSharp.Color color) {
            if (_chart != null) {
                var barChart = _chart.PlotArea.GetFirstChild<BarChart>();
                if (barChart != null) {
                    BarChartSeries barChartSeries = WordBarChart.AddBarChartSeries(this._index, name, color, this.Categories, new List<int>() { values });
                    barChart.Append(barChartSeries);
                }
            }
        }
        public void AddChartBar<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            if (_chart != null) {
                var barChart = _chart.PlotArea.GetFirstChild<BarChart>();
                if (barChart != null) {
                    BarChartSeries barChartSeries = WordBarChart.AddBarChartSeries(this._index, name, color, this.Categories, values);
                    barChart.Append(barChartSeries);
                }
            }
        }

        public void AddChartBar(string name, int[] values, SixLabors.ImageSharp.Color color) {
            if (_chart != null) {
                var barChart = _chart.PlotArea.GetFirstChild<BarChart>();
                if (barChart != null) {
                    BarChartSeries barChartSeries = WordBarChart.AddBarChartSeries(this._index, name, color, this.Categories, values.ToList());
                    barChart.Append(barChartSeries);
                }
            }
        }

        public void AddChartArea<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            if (_chart != null) {
                var barChart = _chart.PlotArea.GetFirstChild<AreaChart>();
                if (barChart != null) {
                    AreaChartSeries areaChartSeries = WordAreaChart.AddAreaChartSeries(this._index, name, color, this.Categories, values);
                    barChart.Append(areaChartSeries);
                }
            }
        }

        public void AddChartArea<T>(string name, int[] values, SixLabors.ImageSharp.Color color) {
            if (_chart != null) {
                var barChart = _chart.PlotArea.GetFirstChild<AreaChart>();
                if (barChart != null) {
                    AreaChartSeries areaChartSeries = WordAreaChart.AddAreaChartSeries(this._index, name, color, this.Categories, values.ToList());
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
