using DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word {
    public partial class WordChart {
        public BarGroupingValues? BarGrouping {
            get {
                var chart = _chartPart.ChartSpace.GetFirstChild<Chart>();
                if (chart != null) {
                    var barChart = chart.PlotArea.GetFirstChild<BarChart>();
                    if (barChart != null) {
                        return barChart.BarGrouping.Val;
                    }
                }

                return null;
            }
            set {
                var chart = _chartPart.ChartSpace.GetFirstChild<Chart>();
                if (chart != null) {
                    var barChart = chart.PlotArea.GetFirstChild<BarChart>();
                    if (barChart != null) {
                        if (barChart.BarGrouping != null) {
                            barChart.BarGrouping.Val = value;
                        }
                    }
                }
            }
        }
        public BarDirectionValues? BarDirection {
            get {
                var chart = _chartPart.ChartSpace.GetFirstChild<Chart>();
                if (chart != null) {
                    var barChart = chart.PlotArea.GetFirstChild<BarChart>();
                    if (barChart != null) {
                        return barChart.BarDirection.Val;
                    }
                }

                return null;
            }
            set {
                var chart = _chartPart.ChartSpace.GetFirstChild<Chart>();
                if (chart != null) {
                    var barChart = chart.PlotArea.GetFirstChild<BarChart>();
                    if (barChart != null) {
                        if (barChart.BarDirection != null) {
                            barChart.BarDirection.Val = value;
                        }
                    }
                }
            }
        }

        public bool RoundedCorners {
            get {
                var roundedCorners = _chartPart.ChartSpace.GetFirstChild<RoundedCorners>();
                if (roundedCorners != null) {
                    return roundedCorners.Val;
                }

                return true;
            }
            set {
                var roundedCorners = _chartPart.ChartSpace.GetFirstChild<RoundedCorners>();
                if (roundedCorners == null) {
                    roundedCorners = new RoundedCorners() { Val = value };
                }
                roundedCorners.Val = value;

            }
        }

        public List<string> Categories { get; set; }

        public List<int> Values { get; set; } = new List<int>();
    }
}
