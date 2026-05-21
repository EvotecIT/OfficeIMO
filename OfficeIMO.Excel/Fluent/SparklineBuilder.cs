using DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent builder for worksheet sparklines.
    /// </summary>
    public sealed class SparklineBuilder {
        private readonly ExcelSheet _sheet;
        private readonly string _dataRange;
        private SparklineTypeValues _type = SparklineTypeValues.Line;
        private bool _displayMarkers;
        private bool _displayHighLow;
        private bool _displayFirstLast;
        private bool _displayNegative;
        private bool _displayAxis;
        private string? _seriesColor;
        private string? _axisColor;
        private string? _negativeColor;
        private string? _markersColor;
        private string? _highColor;
        private string? _lowColor;
        private string? _firstColor;
        private string? _lastColor;

        internal SparklineBuilder(ExcelSheet sheet, string dataRange) {
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            _dataRange = string.IsNullOrWhiteSpace(dataRange)
                ? throw new ArgumentNullException(nameof(dataRange))
                : dataRange;
        }

        /// <summary>Uses a specific sparkline type.</summary>
        public SparklineBuilder Type(SparklineTypeValues type) {
            _type = type;
            return this;
        }

        /// <summary>Uses line sparklines.</summary>
        public SparklineBuilder Line() => Type(SparklineTypeValues.Line);

        /// <summary>Uses column sparklines.</summary>
        public SparklineBuilder Column() => Type(SparklineTypeValues.Column);

        /// <summary>Uses win/loss sparklines.</summary>
        public SparklineBuilder WinLoss() => Type(SparklineTypeValues.Stacked);

        /// <summary>Shows or hides point markers.</summary>
        public SparklineBuilder Markers(bool show = true) {
            _displayMarkers = show;
            return this;
        }

        /// <summary>Shows or hides high and low point markers.</summary>
        public SparklineBuilder HighLow(bool show = true) {
            _displayHighLow = show;
            return this;
        }

        /// <summary>Shows or hides first and last point markers.</summary>
        public SparklineBuilder FirstLast(bool show = true) {
            _displayFirstLast = show;
            return this;
        }

        /// <summary>Shows or hides negative point markers.</summary>
        public SparklineBuilder Negative(bool show = true) {
            _displayNegative = show;
            return this;
        }

        /// <summary>Shows or hides the horizontal axis.</summary>
        public SparklineBuilder Axis(bool show = true) {
            _displayAxis = show;
            return this;
        }

        /// <summary>Sets the main sparkline series color.</summary>
        public SparklineBuilder Color(string color) {
            _seriesColor = RequireColor(color);
            return this;
        }

        /// <summary>Sets optional sparkline colors.</summary>
        public SparklineBuilder Colors(
            string? series = null,
            string? axis = null,
            string? negative = null,
            string? markers = null,
            string? high = null,
            string? low = null,
            string? first = null,
            string? last = null) {
            _seriesColor = NormalizeOptionalColor(series);
            _axisColor = NormalizeOptionalColor(axis);
            _negativeColor = NormalizeOptionalColor(negative);
            _markersColor = NormalizeOptionalColor(markers);
            _highColor = NormalizeOptionalColor(high);
            _lowColor = NormalizeOptionalColor(low);
            _firstColor = NormalizeOptionalColor(first);
            _lastColor = NormalizeOptionalColor(last);
            return this;
        }

        /// <summary>Creates the sparklines at the given location range.</summary>
        public ExcelSheet At(string locationRange) {
            _sheet.AddSparklines(
                _dataRange,
                locationRange,
                _type,
                _displayMarkers,
                _displayHighLow,
                _displayFirstLast,
                _displayNegative,
                _displayAxis,
                _seriesColor,
                _axisColor,
                _negativeColor,
                _markersColor,
                _highColor,
                _lowColor,
                _firstColor,
                _lastColor);
            return _sheet;
        }

        private static string? NormalizeOptionalColor(string? color)
            => string.IsNullOrWhiteSpace(color) ? null : color;

        private static string RequireColor(string color)
            => string.IsNullOrWhiteSpace(color) ? throw new ArgumentNullException(nameof(color)) : color;
    }
}
