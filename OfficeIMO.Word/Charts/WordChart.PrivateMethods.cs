using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word {
    public partial class WordChart {
        /// <summary>
        /// The current index for values
        /// </summary>
        private uint _currentIndexValues = 0;
        /// <summary>
        /// The current index for categories
        /// </summary>
        private UInt32Value _currentIndexCategory = 0;

        internal CategoryAxisData InitializeCategoryAxisData() {
            var pieChartSeries = InitializePieChartSeries();
            CategoryAxisData categoryAxis = pieChartSeries?.GetFirstChild<CategoryAxisData>();
            // If CategoryAxisData does not exist, create it
            if (categoryAxis == null) {
                categoryAxis = new CategoryAxisData();
                StringLiteral stringLiteral = new StringLiteral();
                categoryAxis.Append(stringLiteral);
            }
            return categoryAxis;
        }


        internal NumberLiteral InitializeNumberLiteral() {
            NumberLiteral literal = _chart?.PlotArea?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>()?.GetFirstChild<Values>()?.GetFirstChild<NumberLiteral>();
            // If NumberLiteral does not exist, create it
            if (literal == null) {
                literal = new NumberLiteral();
                FormatCode format = new FormatCode() { Text = "General" };
                literal.Append(format);
            }
            return literal;
        }

        internal Values InitializeValues() {
            var pieChartSeries = InitializePieChartSeries();
            Values values = pieChartSeries?.GetFirstChild<Values>() ?? new Values() { NumberLiteral = InitializeNumberLiteral() };
            return values;
        }

        internal PieChartSeries InitializePieChartSeries() {
            if (_chart != null) {
                var pieChart = _chart.PlotArea.GetFirstChild<PieChart>();
                if (pieChart != null) {
                    var pieChartSeries = pieChart.GetFirstChild<PieChartSeries>();
                    if (pieChartSeries == null) {
                        pieChartSeries = WordPieChart.CreatePieChartSeries(_index, "Title?");
                        pieChart.Append(pieChartSeries);

                    }
                    return pieChartSeries;
                }
            }
            return null;
        }

        internal static PieChartSeries CreatePieChartSeries(UInt32Value index, string series) {
            PieChartSeries pieChartSeries1 = new PieChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index index1 = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = index };

            Order order1 = new Order() { Val = index };
            SeriesText seriesText1 = new SeriesText();

            var stringReference1 = AddSeries(0, series);
            seriesText1.Append(stringReference1);

            InvertIfNegative invertIfNegative1 = new InvertIfNegative();
            pieChartSeries1.Append(index1);
            pieChartSeries1.Append(order1);
            pieChartSeries1.Append(seriesText1);
            pieChartSeries1.Append(invertIfNegative1);
            return pieChartSeries1;
        }

        /// <summary>
        /// Adds single category to charts
        /// </summary>
        /// <param name="category">The category.</param>
        internal void AddSingleCategory(string category) {
            var pieChartSeries = InitializePieChartSeries();

            CategoryAxisData categoryAxis = InitializeCategoryAxisData();

            StringLiteral stringLiteral = categoryAxis.GetFirstChild<StringLiteral>();
            // If StringLiteral does not exist, create it
            if (stringLiteral == null) {
                stringLiteral = new StringLiteral();
                categoryAxis.Append(stringLiteral);
            }
            stringLiteral.Append(new StringPoint() { Index = _currentIndexCategory, NumericValue = new DocumentFormat.OpenXml.Drawing.Charts.NumericValue() { Text = category } });
            // Update the PointCount
            PointCount pointCount = stringLiteral.GetFirstChild<PointCount>();
            if (pointCount != null) {
                pointCount.Val = _currentIndexCategory + 1;
            } else {
                stringLiteral.InsertAt(new PointCount() { Val = 1 }, 0);
            }
            // Increment the current index
            _currentIndexCategory++;

            if (!pieChartSeries.Elements<CategoryAxisData>().Any()) {
                pieChartSeries.Append(categoryAxis);
            }
        }

        /// <summary>
        /// Adds the single value to charts
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data">The data.</param>
        internal void AddSingleValue<T>(T data) {
            // Initialize the PieChartSeries
            var pieChartSeries = InitializePieChartSeries();
            // Initialize the Values
            Values values = InitializeValues();
            // Initialize the NumberLiteral
            NumberLiteral literal = values.GetFirstChild<NumberLiteral>() ?? InitializeNumberLiteral();
            literal.Append(new NumericPoint() { Index = _currentIndexValues, NumericValue = new NumericValue() { Text = data.ToString() } });
            // Update the PointCount
            PointCount pointCount = literal.GetFirstChild<PointCount>();
            if (pointCount != null) {
                pointCount.Val = _currentIndexValues + 1;
            } else {
                literal.InsertAt(new PointCount() { Val = 1 }, 0);
            }
            // Increment the current index
            _currentIndexValues++;
            // add values to the series if it does not exist
            if (!pieChartSeries.Elements<Values>().Any()) {
                pieChartSeries.Append(values);
            }
        }

    }
}
