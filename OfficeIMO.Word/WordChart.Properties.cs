using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordChart {
        private const long EnglishMetricUnitsPerInch = 914400;
        private const long PixelsPerInch = 96;
        private readonly WordDocument _document;
        private WordParagraph _paragraph;
        private ChartPart _chartPart {
            get {
                if (_drawing == null) {
                    return null;
                } else {
                    var id = _drawing.Inline.Graphic.GraphicData.GetFirstChild<ChartReference>().Id;
                    return (ChartPart)_document._wordprocessingDocument.MainDocumentPart.GetPartById(id);
                }
            }
        }
        private Drawing _drawing;
        private Chart _chart;
        /// <summary>
        /// The current index for values
        /// </summary>
        private uint _currentIndexValues = 0;
        /// <summary>
        /// The current index for categories
        /// </summary>
        private UInt32Value _currentIndexCategory = 0;
        //private string _id => _document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(_chartPart);

        public BarGroupingValues? BarGrouping {
            get {
                if (_chartPart != null) {
                    var chart = _chartPart.ChartSpace.GetFirstChild<Chart>();
                    if (chart != null) {
                        var barChart = chart.PlotArea.GetFirstChild<BarChart>();
                        if (barChart != null) {
                            return barChart.BarGrouping.Val;
                        }
                    }
                }

                return null;
            }
            set {
                if (_chartPart != null) {
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
        }
        public BarDirectionValues? BarDirection {
            get {
                if (_chartPart != null) {
                    var chart = _chartPart.ChartSpace.GetFirstChild<Chart>();
                    if (chart != null) {
                        var barChart = chart.PlotArea.GetFirstChild<BarChart>();
                        if (barChart != null) {
                            return barChart.BarDirection.Val;
                        }
                    }
                }

                return null;
            }
            set {
                if (_chartPart != null) {
                    var chart = _chartPart.ChartSpace.GetFirstChild<Chart>();
                    if (chart != null) {
                        if (chart.PlotArea != null) {
                            var barChart = chart.PlotArea.GetFirstChild<BarChart>();
                            if (barChart != null) {
                                if (barChart.BarDirection != null) {
                                    barChart.BarDirection.Val = value;
                                }
                            }
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

        private List<string> Categories { get; set; }

        /// <summary>
        /// Holds the title of the chart until we can add it to the chart
        ///
        /// Note: the title can't really be added to a chart before we know what type of chart it is
        /// Since we don't know what type of chart it is until we add the first Pie, Bar, Line or Area
        /// we need to wait until then to add the title
        /// This is why we have a separate property for the title, but the method to add the title is in the WordChart class
        /// </summary>
        private string PrivateTitle { get; set; }

        /// <summary>
        /// Get or set the title of the chart
        /// </summary>
        public string Title {
            get {
                if (_chart != null && _chart.Title != null) {
                    var text = _chart.Title.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault();
                    if (text != null) {
                        return text.Text;
                    }
                }
                if (PrivateTitle != null) {
                    return PrivateTitle;
                }
                return null;
            }
            set {
                if (_chart != null) {
                    SetTitle(value);
                }
                PrivateTitle = value;
            }
        }
    }
}
