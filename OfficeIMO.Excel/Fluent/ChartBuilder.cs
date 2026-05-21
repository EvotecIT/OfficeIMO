namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent builder for creating charts from an A1 range or table.
    /// </summary>
    public sealed class ChartBuilder {
        private readonly ExcelSheet _sheet;
        private readonly string _source;
        private readonly bool _isTableSource;
        private ExcelChartType _type = ExcelChartType.ColumnClustered;
        private bool _hasHeaders = true;
        private bool _includeCachedData = true;
        private string? _title;
        private int _widthPixels = 640;
        private int _heightPixels = 360;

        internal ChartBuilder(ExcelSheet sheet, string source, bool isTableSource) {
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            _source = string.IsNullOrWhiteSpace(source)
                ? throw new ArgumentNullException(nameof(source))
                : source;
            _isTableSource = isTableSource;
        }

        /// <summary>Uses a specific chart type.</summary>
        public ChartBuilder Type(ExcelChartType type) {
            _type = type;
            return this;
        }

        /// <summary>Uses a clustered column chart.</summary>
        public ChartBuilder ColumnClustered() => Type(ExcelChartType.ColumnClustered);

        /// <summary>Uses a stacked column chart.</summary>
        public ChartBuilder ColumnStacked() => Type(ExcelChartType.ColumnStacked);

        /// <summary>Uses a clustered bar chart.</summary>
        public ChartBuilder BarClustered() => Type(ExcelChartType.BarClustered);

        /// <summary>Uses a stacked bar chart.</summary>
        public ChartBuilder BarStacked() => Type(ExcelChartType.BarStacked);

        /// <summary>Uses a line chart.</summary>
        public ChartBuilder Line() => Type(ExcelChartType.Line);

        /// <summary>Uses an area chart.</summary>
        public ChartBuilder Area() => Type(ExcelChartType.Area);

        /// <summary>Uses a pie chart.</summary>
        public ChartBuilder Pie() => Type(ExcelChartType.Pie);

        /// <summary>Uses a doughnut chart.</summary>
        public ChartBuilder Doughnut() => Type(ExcelChartType.Doughnut);

        /// <summary>Uses a scatter chart. Category values must be numeric.</summary>
        public ChartBuilder Scatter() => Type(ExcelChartType.Scatter);

        /// <summary>Uses defaults suitable for a time-series revenue or volume trend.</summary>
        public ChartBuilder RevenueTrend(string title = "Revenue Trend", int widthPixels = 720, int heightPixels = 320) {
            return Recipe(ExcelChartType.Line, title, widthPixels, heightPixels);
        }

        /// <summary>Uses defaults suitable for a status, category, or allocation breakdown.</summary>
        public ChartBuilder StatusBreakdown(string title = "Status Breakdown", int widthPixels = 520, int heightPixels = 320) {
            return Recipe(ExcelChartType.Doughnut, title, widthPixels, heightPixels);
        }

        /// <summary>Uses defaults suitable for ranking the largest items in a compact dashboard.</summary>
        public ChartBuilder TopNBar(string title = "Top Items", int widthPixels = 640, int heightPixels = 360) {
            return Recipe(ExcelChartType.BarClustered, title, widthPixels, heightPixels);
        }

        /// <summary>Uses defaults suitable for positive/negative variance comparisons.</summary>
        public ChartBuilder VarianceColumns(string title = "Variance", int widthPixels = 640, int heightPixels = 360) {
            return Recipe(ExcelChartType.ColumnClustered, title, widthPixels, heightPixels);
        }

        /// <summary>Sets the chart title.</summary>
        public ChartBuilder Title(string title) {
            _title = string.IsNullOrWhiteSpace(title) ? throw new ArgumentNullException(nameof(title)) : title;
            return this;
        }

        /// <summary>Controls whether the first row of a range contains headers.</summary>
        public ChartBuilder Headers(bool hasHeaders = true) {
            _hasHeaders = hasHeaders;
            return this;
        }

        /// <summary>Controls whether cached chart data is written.</summary>
        public ChartBuilder CachedData(bool includeCachedData = true) {
            _includeCachedData = includeCachedData;
            return this;
        }

        /// <summary>Sets the chart dimensions in pixels.</summary>
        public ChartBuilder Size(int widthPixels, int heightPixels) {
            if (widthPixels <= 0) throw new ArgumentOutOfRangeException(nameof(widthPixels));
            if (heightPixels <= 0) throw new ArgumentOutOfRangeException(nameof(heightPixels));
            _widthPixels = widthPixels;
            _heightPixels = heightPixels;
            return this;
        }

        /// <summary>Creates the chart at the given worksheet coordinates.</summary>
        public ExcelChart At(int row, int column) {
            if (_isTableSource) {
                return _sheet.AddChartFromTable(_source, row, column, _widthPixels, _heightPixels, _type, _title, _includeCachedData);
            }

            return _sheet.AddChartFromRange(_source, row, column, _widthPixels, _heightPixels, _type, _hasHeaders, _title, _includeCachedData);
        }

        private ChartBuilder Recipe(ExcelChartType type, string title, int widthPixels, int heightPixels) {
            Type(type);
            Title(title);
            Size(widthPixels, heightPixels);
            return this;
        }
    }
}
