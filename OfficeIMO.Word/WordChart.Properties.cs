using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    // Partial class containing property definitions and helper fields for
    // WordChart.
    public partial class WordChart {
        /// <summary>
        /// Built‑in color palettes for charts.
        /// Used by <see cref="WordChart.ApplyPalette(WordChartPalette, bool, bool, bool, System.Collections.Generic.Dictionary{string, OfficeIMO.OfficeColor}?)"/>
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
        /// <summary>The chart's <c>w:drawing</c>.</summary>
        private OfficeXmlNode _drawing;
        /// <summary>The chart's <c>c:chart</c>, or none until the first series creates it.</summary>
        private OfficeXmlNode _chart;

        /// <summary>
        /// The current index for values
        /// </summary>
        private uint _currentIndexValues = 0;
        /// <summary>
        /// The current index for categories
        /// </summary>
        private uint _currentIndexCategory = 0;
        private string? _xAxisTitle;
        private string? _yAxisTitle;
        private string _axisTitleFontName = OfficeTheme.Default.ChartStyle.AxisFont.Family;
        private int _axisTitleFontSize = 1100;
        private OfficeIMO.OfficeColor _axisTitleColor = OfficeIMO.OfficeColor.Black;

        /// <summary>The first <c>c:barChart</c> of the chart part's chart, or none.</summary>
        private OfficeXmlNode SavedBarChart {
            get {
                OfficeXmlNode space = ChartSpaceNode;
                if (space.IsNone) return default;
                OfficeXmlNode plot = PlotAreaOf(space.Child(NsC, "chart"));
                return plot.IsNone ? default : plot.Child(NsC, "barChart");
            }
        }

        /// <summary>
        /// Gets or sets the bar grouping mode for bar charts.
        /// </summary>
        public WordChartBarGrouping? BarGrouping {
            get {
                OfficeXmlNode grouping = SavedBarChart.IsNone ? default : SavedBarChart.Child(NsC, "grouping");
                if (grouping.IsNone || grouping.Val() is not string token) return null;
                return WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordChartBarGrouping value)
                    ? value : throw new FormatException("The bar grouping '" + token + "' is not a valid value.");
            }
            set {
                OfficeXmlNode barChart = SavedBarChart;
                if (barChart.IsNone) return;
                if (value.HasValue) {
                    OfficeXmlNode grouping = barChart.Child(NsC, "grouping");
                    if (grouping.IsNone) grouping = OfficeChartXml.SetOrdered(barChart, NewC(barChart, "grouping"), CBarChartView.Order);
                    grouping.SetAttribute("val", value.Value.ToOpenXmlToken());
                    string overlapValue = value.Value is WordChartBarGrouping.Stacked or WordChartBarGrouping.PercentStacked ? "100" : "0";
                    OfficeXmlNode overlap = barChart.Child(NsC, "overlap");
                    if (overlap.IsNone) {
                        OfficeChartXml.SetOrdered(barChart, NewC(barChart, "overlap", overlapValue), CBarChartView.Order);
                    } else {
                        overlap.SetAttribute("val", overlapValue);
                    }
                } else {
                    barChart.RemoveChildren(NsC, "grouping");
                    OfficeXmlNode overlap = barChart.Child(NsC, "overlap");
                    if (!overlap.IsNone) overlap.SetAttribute("val", "0");
                }
            }
        }
        /// <summary>
        /// Gets or sets the bar direction (row or column) for bar charts.
        /// </summary>
        public WordChartBarDirection? BarDirection {
            get {
                OfficeXmlNode direction = SavedBarChart.IsNone ? default : SavedBarChart.Child(NsC, "barDir");
                if (direction.IsNone || direction.Val() is not string token) return null;
                return WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordChartBarDirection value)
                    ? value : throw new FormatException("The bar direction '" + token + "' is not a valid value.");
            }
            set {
                OfficeXmlNode barChart = SavedBarChart;
                if (barChart.IsNone) return;
                if (value.HasValue) {
                    OfficeXmlNode direction = barChart.Child(NsC, "barDir");
                    if (direction.IsNone) direction = OfficeChartXml.SetOrdered(barChart, NewC(barChart, "barDir"), CBarChartView.Order);
                    direction.SetAttribute("val", value.Value.ToOpenXmlToken());
                    SetSingleBarAxisPositions(barChart.Parent, value.Value == WordChartBarDirection.Bar ? "l" : "b",
                        value.Value == WordChartBarDirection.Bar ? "b" : "l");
                } else {
                    barChart.RemoveChildren(NsC, "barDir");
                    SetSingleBarAxisPositions(barChart.Parent, "b", "l");
                }
            }
        }

        /// <summary>
        /// Places the category and value axes of a plot area holding one chart group on one pair of axes; any other
        /// plot area keeps its axes where they are.
        /// </summary>
        private static void SetSingleBarAxisPositions(OfficeXmlNode plot, string categoryPosition, string valuePosition) {
            int groups = 0;
            foreach (OfficeXmlNode child in plot.Children()) {
                if (IsChartGroup(child)) groups++;
            }
            if (groups != 1 || plot.Children(NsC, "catAx").Count != 1 || plot.Children(NsC, "valAx").Count != 1) return;
            OfficeXmlNode category = plot.Child(NsC, "catAx").Child(NsC, "axPos");
            OfficeXmlNode numeric = plot.Child(NsC, "valAx").Child(NsC, "axPos");
            if (category.IsNone || numeric.IsNone) return;
            category.SetAttribute("val", categoryPosition);
            numeric.SetAttribute("val", valuePosition);
        }

        /// <summary>
        /// Gets or sets whether the chart frame uses rounded corners.
        /// </summary>
        public bool RoundedCorners {
            get {
                OfficeXmlNode space = ChartSpaceNode;
                OfficeXmlNode roundedCorners = space.IsNone ? default : space.Child(NsC, "roundedCorners");
                return (roundedCorners.IsNone ? null : roundedCorners.BoolVal()) ?? true;
            }
            set {
                OfficeXmlNode space = ChartSpaceNode;
                if (space.IsNone) return;
                OfficeXmlNode roundedCorners = space.Child(NsC, "roundedCorners");
                if (roundedCorners.IsNone) {
                    // c:roundedCorners follows c:date1904 and c:lang in the chart-space sequence.
                    OfficeXmlNode previous = space.Child(NsC, "lang");
                    if (previous.IsNone) previous = space.Child(NsC, "date1904");
                    roundedCorners = previous.IsNone
                        ? space.InsertBefore(NewC(space, "roundedCorners"), space.FirstChild)
                        : previous.InsertAfterSelf(NewC(space, "roundedCorners"));
                }
                roundedCorners.SetAttribute("val", OfficeChartXml.Text(value));
            }
        }

        private List<string>? _categories;

        /// <summary>Categories used by series added later; a saved chart starts with its first series' categories.</summary>
        private List<string> Categories {
            get => _categories ??= ReadSavedCategories();
            set => _categories = value;
        }

        private List<string> ReadSavedCategories() {
            // The first series in the plot area's chart groups that has categories.
            OfficeXmlNode plot = PlotAreaOf(_chart);
            for (OfficeXmlNode group = plot.IsNone ? default : plot.FirstChild; !group.IsNone; group = group.NextSibling) {
                if (group.NamespaceId != NsC || !IsChartGroup(group)) continue;
                foreach (OfficeXmlNode series in group.Children(NsC, "ser")) {
                    OfficeXmlNode categories = series.Child(NsC, "cat");
                    if (!categories.IsNone) return ReadCachedStrings(categories).ToList();
                }
            }
            return new List<string>();
        }

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
                string? saved = !_chart.IsNone ? ReadTitle(_chart) : null;
                if (saved != null) {
                    return saved;
                }
                if (!string.IsNullOrEmpty(PrivateTitle)) {
                    return PrivateTitle;
                }
                return null;
            }
            set {
                if (!_chart.IsNone && value != null) {
                    SetTitle(value);
                }
                PrivateTitle = value ?? string.Empty;
            }
        }
    }
}
