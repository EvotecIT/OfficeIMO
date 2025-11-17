using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Partial class containing property definitions and helper fields for
    /// <see cref="WordChart"/>.
    /// </summary>
    public partial class WordChart {
        /// <summary>
        /// Built‑in color palettes for charts.
        /// Used by <see cref="WordChart.ApplyPalette(WordChartPalette, bool, bool, bool, System.Collections.Generic.Dictionary{string, SixLabors.ImageSharp.Color}?)"/>
        /// to style series and pie slices in a consistent, professional way.
        /// </summary>
        public enum WordChartPalette {
            /// <summary>
            /// Professional palette optimized for executive reports — balanced blues/greens/oranges with a neutral gray.
            /// </summary>
            Professional,
            /// <summary>
            /// Soft pastel palette suited for dashboards and slides; low‑contrast tints.
            /// </summary>
            Soft,
            /// <summary>
            /// Monochrome gray scale for print‑friendly output and high legibility.
            /// </summary>
            MonochromeGray,
            /// <summary>
            /// Color‑blind friendly palette (Okabe–Ito), designed for strong category separation.
            /// </summary>
            ColorBlindSafe
        }

        private const long EnglishMetricUnitsPerInch = 914400;
        private const long PixelsPerInch = 96;
        private readonly WordDocument _document;
        private WordParagraph? _paragraph;
        private ChartPart? _chartPart {
            get {
                if (_drawing == null) {
                    return null;
                }
                var chartRef = _drawing.Inline?.Graphic?.GraphicData?.GetFirstChild<ChartReference>();
                var id = chartRef?.Id?.Value;
                return id != null ? (ChartPart?)_document._wordprocessingDocument.MainDocumentPart!.GetPartById(id) : null;
            }
        }
        private Drawing? _drawing;
        private Chart? _chart;
        /// <summary>
        /// The current index for values
        /// </summary>
        private uint _currentIndexValues = 0;
        /// <summary>
        /// The current index for categories
        /// </summary>
        private UInt32Value _currentIndexCategory = 0;
        private string? _xAxisTitle;
        private string? _yAxisTitle;
        private string _axisTitleFontName = "Calibri";
        private int _axisTitleFontSize = 11;
        private SixLabors.ImageSharp.Color _axisTitleColor = SixLabors.ImageSharp.Color.Black;
        //private string _id => _document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(_chartPart);

        /// <summary>
        /// Gets or sets the bar grouping mode for bar charts.
        /// </summary>
        public BarGroupingValues? BarGrouping {
            get {
                if (_chartPart != null) {
                    var chart = _chartPart.ChartSpace.GetFirstChild<Chart>();
                    var barChart = chart?.PlotArea?.GetFirstChild<BarChart>();
                    if (barChart?.BarGrouping != null) {
                        return barChart.BarGrouping.Val?.Value;
                    }
                }

                return null;
            }
            set {
                if (_chartPart != null) {
                    var chart = _chartPart.ChartSpace.GetFirstChild<Chart>();
                    var barChart = chart?.PlotArea?.GetFirstChild<BarChart>();
                    if (barChart != null) {
                        if (value.HasValue) {
                            (barChart.BarGrouping ??= new BarGrouping()).Val = value.Value;
                        } else {
                            barChart.BarGrouping = null;
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Gets or sets the bar direction (row or column) for bar charts.
        /// </summary>
        public BarDirectionValues? BarDirection {
            get {
                if (_chartPart != null) {
                    var chart = _chartPart.ChartSpace.GetFirstChild<Chart>();
                    var barChart = chart?.PlotArea?.GetFirstChild<BarChart>();
                    if (barChart?.BarDirection != null) {
                        return barChart.BarDirection.Val?.Value;
                    }
                }

                return null;
            }
            set {
                if (_chartPart != null) {
                    var chart = _chartPart.ChartSpace.GetFirstChild<Chart>();
                    var barChart = chart?.PlotArea?.GetFirstChild<BarChart>();
                    if (barChart != null) {
                        if (value.HasValue) {
                            (barChart.BarDirection ??= new BarDirection()).Val = value.Value;
                        } else {
                            barChart.BarDirection = null;
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Gets or sets whether the chart frame uses rounded corners.
        /// </summary>
        public bool RoundedCorners {
            get {
                var roundedCorners = _chartPart?.ChartSpace.GetFirstChild<RoundedCorners>();
                if (roundedCorners?.Val != null) {
                    return roundedCorners.Val;
                }

                return true;
            }
            set {
                if (_chartPart == null) return;
                var roundedCorners = _chartPart.ChartSpace.GetFirstChild<RoundedCorners>();
                if (roundedCorners == null) {
                    roundedCorners = new RoundedCorners() { Val = value };
                }
                roundedCorners.Val = value;

            }
        }

        private List<string> Categories { get; set; } = new();

        /// <summary>
        /// Holds the title of the chart until we can add it to the chart
        ///
        /// Note: the title can't really be added to a chart before we know what type of chart it is
        /// Since we don't know what type of chart it is until we add the first Pie, Bar, Line or Area
        /// we need to wait until then to add the title
        /// This is why we have a separate property for the title, but the method to add the title is in the WordChart class
        /// </summary>
        private string PrivateTitle { get; set; } = string.Empty;

        /// <summary>
        /// Get or set the title of the chart
        /// </summary>
        public string? Title {
            get {
                if (_chart?.Title != null) {
                    var text = _chart.Title.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault();
                    if (text != null) {
                        return text.Text;
                    }
                }
                if (!string.IsNullOrEmpty(PrivateTitle)) {
                    return PrivateTitle;
                }
                return null;
            }
            set {
                if (_chart != null && value != null) {
                    SetTitle(value);
                }
                PrivateTitle = value ?? string.Empty;
            }
        }
    }
}
