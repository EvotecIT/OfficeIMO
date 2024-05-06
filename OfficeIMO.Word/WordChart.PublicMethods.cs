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
            // since the title may have changed, we need to update it
            UpdateTitle();
            return this;
        }

        public void AddChartLine<T>(string name, int[] values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsLine();
            if (InternalChart != null) {
                var lineChart = InternalChart.PlotArea.GetFirstChild<LineChart>();
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
            var lineChart = InternalChart.PlotArea.GetFirstChild<LineChart>();
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
            var barChart = InternalChart.PlotArea.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, new List<int>() { values });
                barChart.Append(barChartSeries);
            }
        }

        public void AddBar<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar();
            var barChart = InternalChart.PlotArea.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, values);
                barChart.Append(barChartSeries);
            }
        }

        public void AddBar(string name, int[] values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsBar();
            var barChart = InternalChart.PlotArea.GetFirstChild<BarChart>();
            if (barChart != null) {
                BarChartSeries barChartSeries = AddBarChartSeries(this._index, name, color, this.Categories, values.ToList());
                barChart.Append(barChartSeries);
            }
        }

        public void AddArea<T>(string name, List<T> values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsArea();
            if (InternalChart != null) {
                var barChart = InternalChart.PlotArea.GetFirstChild<AreaChart>();
                if (barChart != null) {
                    AreaChartSeries areaChartSeries = AddAreaChartSeries(this._index, name, color, this.Categories, values);
                    barChart.Append(areaChartSeries);
                }
            }
        }

        public void AddArea<T>(string name, int[] values, SixLabors.ImageSharp.Color color) {
            EnsureChartExistsArea();
            if (InternalChart != null) {
                var barChart = InternalChart.PlotArea.GetFirstChild<AreaChart>();
                if (barChart != null) {
                    AreaChartSeries areaChartSeries = AddAreaChartSeries(this._index, name, color, this.Categories, values.ToList());
                    barChart.Append(areaChartSeries);
                }
            }
        }

        public void AddLegend(LegendPositionValues legendPosition) {
            if (InternalChart != null) {
                Legend legend = new Legend();
                LegendPosition postion = new LegendPosition() { Val = legendPosition };
                Overlay overlay = new Overlay() { Val = false };
                legend.Append(postion);
                legend.Append(overlay);
                InternalChart.Append(legend);
            }
        }

        public View3D GenerateView3D() {
            View3D view3D1 = new View3D();
            RotateX rotateX1 = new RotateX() { Val = 15 };
            RotateY rotateY1 = new RotateY() { Val = (UInt16Value)20U };
            RightAngleAxes rightAngleAxes1 = new RightAngleAxes() { Val = false };

            view3D1.Append(rotateX1);
            view3D1.Append(rotateY1);
            view3D1.Append(rightAngleAxes1);
            return view3D1;
        }

        public Floor GenerateFloor() {
            Floor floor1 = new Floor();
            Thickness thickness1 = new Thickness() { Val = 0 };

            floor1.Append(thickness1);
            return floor1;
        }


        public SideWall GenerateSideWall() {
            SideWall sideWall1 = new SideWall();
            Thickness thickness1 = new Thickness() { Val = 0 };

            sideWall1.Append(thickness1);
            return sideWall1;
        }

        public BackWall GenerateBackWall() {
            BackWall backWall1 = new BackWall();
            Thickness thickness1 = new Thickness() { Val = 0 };

            backWall1.Append(thickness1);
            return backWall1;
        }



        public PlotVisibleOnly GeneratePlotVisibleOnly() {
            PlotVisibleOnly plotVisibleOnly1 = new PlotVisibleOnly() { Val = true };
            return plotVisibleOnly1;
        }


        public DisplayBlanksAs GenerateDisplayBlanksAs() {
            DisplayBlanksAs displayBlanksAs1 = new DisplayBlanksAs() { Val = DisplayBlanksAsValues.Gap };
            return displayBlanksAs1;
        }


        public ShowDataLabelsOverMaximum GenerateShowDataLabelsOverMaximum() {
            ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new ShowDataLabelsOverMaximum() { Val = false };
            return showDataLabelsOverMaximum1;
        }
    }
}
