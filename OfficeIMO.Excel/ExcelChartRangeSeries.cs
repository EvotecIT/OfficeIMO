using System.Globalization;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using ChartIndex = DocumentFormat.OpenXml.Drawing.Charts.Index;

namespace OfficeIMO.Excel {
    /// <summary>Category-chart series that reference their own worksheet ranges, validated before any part is written.</summary>
    internal sealed class ExcelChartRangeSeries {
        private readonly IReadOnlyList<ExcelChartSeriesRange> _source;
        private readonly Vector _categories;
        private readonly Vector[] _values;

        private ExcelChartRangeSeries(IReadOnlyList<ExcelChartSeriesRange> source, Vector categories, Vector[] values) {
            _source = source;
            _categories = categories;
            _values = values;
        }

        internal int PointCount => _categories.Count;

        internal static ExcelChartRangeSeries Resolve(ExcelDocument document, string defaultSheetName, IReadOnlyList<ExcelChartSeriesRange> ranges) {
            if (ranges.Count == 0) throw new ArgumentException("At least one series range is required.", nameof(ranges));
            Vector? categories = null;
            var values = new Vector[ranges.Count];
            for (int i = 0; i < ranges.Count; i++) {
                ExcelChartSeriesRange item = ranges[i] ?? throw new ArgumentNullException(nameof(ranges), "Series ranges cannot contain null.");
                if (item.BubbleSizeRangeA1 != null) {
                    throw new ArgumentException("Bubble size ranges apply only to bubble charts.", nameof(ranges));
                }

                Vector itemCategories = Vector.Parse(document, defaultSheetName, item.XRangeA1);
                if (categories == null) {
                    categories = itemCategories;
                } else if (!categories.SameCells(itemCategories)) {
                    throw new ArgumentException("Every series of a category chart must use the same category range.", nameof(ranges));
                }

                values[i] = Vector.Parse(document, defaultSheetName, item.YRangeA1);
                if (values[i].Count != categories.Count) {
                    throw new ArgumentException($"Values range '{item.YRangeA1}' has {values[i].Count} cells but the category range has {categories.Count}.", nameof(ranges));
                }
            }

            return new ExcelChartRangeSeries(ranges, categories!, values);
        }

        internal void ValidateCache(ExcelChartData data) {
            if (data.Series.Count != _values.Length || data.Categories.Count != PointCount) {
                throw new ArgumentException("Cached chart data must have one series per range and one category per referenced cell.", nameof(data));
            }
        }

        /// <summary>Reads the snapshot cache from the referenced cells, the same way linked range charts do.</summary>
        internal ExcelChartData ReadCache(ExcelDocument document) {
            ExcelSheet categorySheet = document[_categories.SheetName];
            var categories = new string[PointCount];
            for (int i = 0; i < categories.Length; i++) {
                (int row, int column) = _categories.Cell(i);
                categories[i] = categorySheet.TryGetCellText(row, column, out string? text) ? text ?? string.Empty : string.Empty;
            }

            var series = new List<ExcelChartSeries>(_values.Length);
            for (int s = 0; s < _values.Length; s++) {
                Vector vector = _values[s];
                ExcelSheet sheet = document[vector.SheetName];
                var points = new double[vector.Count];
                for (int i = 0; i < points.Length; i++) {
                    (int row, int column) = vector.Cell(i);
                    points[i] = sheet.TryGetCellText(row, column, out string? raw) &&
                        double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double value) ? value : 0D;
                }

                series.Add(ExcelChartSeries.CreateOwned(NameAt(s, null), points));
            }

            return new ExcelChartData(categories, series);
        }

        internal IReadOnlyList<string> ResolveNames(ExcelChartData? data) {
            var names = new string[_values.Length];
            for (int i = 0; i < names.Length; i++) names[i] = NameAt(i, data);
            return names;
        }

        private string NameAt(int index, ExcelChartData? data) {
            if (!string.IsNullOrWhiteSpace(_source[index].Name)) return _source[index].Name;
            string? cached = data != null && index < data.Series.Count ? data.Series[index].Name : null;
            return string.IsNullOrWhiteSpace(cached) ? $"Series {index + 1}" : cached!;
        }

        /// <summary>Points each category-chart series at its own category and values ranges.</summary>
        internal void ApplyFormulas(ChartPart chartPart) {
            PlotArea? plotArea = chartPart.ChartSpace?.GetFirstChild<Chart>()?.GetFirstChild<PlotArea>();
            if (plotArea == null) return;
            string categoryFormula = _categories.Formula;
            foreach (CategoryAxisData categories in plotArea.Descendants<CategoryAxisData>()) {
                Formula? formula = categories.GetFirstChild<StringReference>()?.Formula ?? categories.GetFirstChild<NumberReference>()?.Formula;
                if (formula != null) formula.Text = categoryFormula;
            }

            foreach (Values values in plotArea.Descendants<Values>()) {
                uint? index = values.Parent?.GetFirstChild<ChartIndex>()?.Val?.Value;
                Formula? formula = values.GetFirstChild<NumberReference>()?.Formula;
                if (formula != null && index.HasValue && index.Value < (uint)_values.Length) {
                    formula.Text = _values[(int)index.Value].Formula;
                }
            }
        }

        /// <summary>A single-row or single-column worksheet range.</summary>
        private sealed class Vector {
            private Vector(string sheetName, int top, int left, int bottom, int right) {
                SheetName = sheetName;
                Top = top;
                Left = left;
                Bottom = bottom;
                Right = right;
            }

            internal string SheetName { get; }
            private int Top { get; }
            private int Left { get; }
            private int Bottom { get; }
            private int Right { get; }
            internal int Count => Math.Max(Bottom - Top, Right - Left) + 1;
            internal string Formula => ExcelChartUtils.BuildSheetQualifiedRange(SheetName, ExcelChartUtils.BuildRangeA1(Top, Left, Bottom, Right));

            internal (int Row, int Column) Cell(int index) => Top == Bottom ? (Top, Left + index) : (Top + index, Left);

            internal bool SameCells(Vector other) =>
                string.Equals(SheetName, other.SheetName, StringComparison.OrdinalIgnoreCase) &&
                Top == other.Top && Left == other.Left && Bottom == other.Bottom && Right == other.Right;

            internal static Vector Parse(ExcelDocument document, string defaultSheetName, string rangeA1) {
                string qualified = ExcelChartUtils.EnsureSheetQualifiedRange(defaultSheetName, rangeA1);
                if (!ExcelChartUtils.TryParseSheetQualifiedRange(qualified, out string sheetName, out string reference)) {
                    throw new ArgumentException($"Invalid range '{rangeA1}'.", nameof(rangeA1));
                }

                int top, left, bottom, right;
                if (reference.Contains(":")) {
                    if (!ExcelA1.TryParseRange(reference, out top, out left, out bottom, out right)) {
                        throw new ArgumentException($"Invalid range '{rangeA1}'.", nameof(rangeA1));
                    }
                } else {
                    (top, left) = ExcelA1.ParseCellRef(reference);
                    bottom = top;
                    right = left;
                }

                if (top <= 0 || left <= 0 || bottom > ExcelA1.MaxRows || right > ExcelA1.MaxColumns) {
                    throw new ArgumentException($"Invalid range '{rangeA1}'.", nameof(rangeA1));
                }
                if (top != bottom && left != right) {
                    throw new ArgumentException($"Range '{rangeA1}' must be a single row or a single column.", nameof(rangeA1));
                }
                if (!document.TryGetSheet(sheetName, out ExcelSheet? sheet) || sheet == null) {
                    throw new ArgumentException($"Range '{rangeA1}' refers to a missing sheet '{sheetName}'.", nameof(rangeA1));
                }

                return new Vector(sheet.Name, top, left, bottom, right);
            }
        }
    }
}
