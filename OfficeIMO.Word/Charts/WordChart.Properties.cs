using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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

        /// <summary>
        /// The current index for values
        /// </summary>
        private uint _currentIndexValues = 0;
        /// <summary>
        /// The current index for categories
        /// </summary>
        private UInt32Value _currentIndexCategory = 0;

        public string Title { get; set; }

        private WordDocument _document;
        private WordParagraph _paragraph;
        private ChartPart _chartPart;
        private Drawing _drawing;
        internal Chart InternalChart;
        private const long EnglishMetricUnitsPerInch = 914400;
        private const long PixelsPerInch = 96;

        //private string _id => _document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(_chartPart);

    }
}
