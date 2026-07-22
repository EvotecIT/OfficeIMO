using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;
using OfficeReferenceSequence = DocumentFormat.OpenXml.Office.Excel.ReferenceSequence;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Helpers for adding sparklines to worksheets.
    /// </summary>
    public partial class ExcelSheet {
        private const int MaximumExpandedSparklines = 10_000;

        /// <summary>
        /// Adds sparklines to the worksheet.
        /// </summary>
        /// <param name="dataRange">A1 range containing the data (e.g., "B2:M2").</param>
        /// <param name="locationRange">A1 range where sparklines will be placed (e.g., "N2:N2").</param>
        /// <param name="type">Sparkline type.</param>
        /// <param name="displayMarkers">Show markers for each data point.</param>
        /// <param name="displayHighLow">Show high/low markers.</param>
        /// <param name="displayFirstLast">Show first/last markers.</param>
        /// <param name="displayNegative">Show negative markers.</param>
        /// <param name="displayAxis">Show axis.</param>
        /// <param name="seriesColor">Sparkline series color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="axisColor">Axis color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="negativeColor">Negative point color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="markersColor">Markers color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="highColor">High point color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="lowColor">Low point color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="firstColor">First point color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="lastColor">Last point color (#RRGGBB or #AARRGGBB).</param>
        public void AddSparklines(
            string dataRange,
            string locationRange,
            SparklineTypeValues type,
            bool displayMarkers = false,
            bool displayHighLow = false,
            bool displayFirstLast = false,
            bool displayNegative = false,
            bool displayAxis = false,
            string? seriesColor = null,
            string? axisColor = null,
            string? negativeColor = null,
            string? markersColor = null,
            string? highColor = null,
            string? lowColor = null,
            string? firstColor = null,
            string? lastColor = null) {
            if (string.IsNullOrWhiteSpace(dataRange)) throw new ArgumentException("DataRange is required.", nameof(dataRange));
            if (string.IsNullOrWhiteSpace(locationRange)) throw new ArgumentException("LocationRange is required.", nameof(locationRange));

            WriteLock(() => {
                var ws = WorksheetRoot;
                var groups = GetOrCreateSparklineGroups(ws);

                var group = new SparklineGroup { Type = type };
                if (displayMarkers) group.Markers = true;
                if (displayHighLow) {
                    group.High = true;
                    group.Low = true;
                }
                if (displayFirstLast) {
                    group.First = true;
                    group.Last = true;
                }
                if (displayNegative) group.Negative = true;
                if (displayAxis) group.DisplayXAxis = true;

                ApplyColor(seriesColor, rgb => group.SeriesColor = new SeriesColor { Rgb = rgb });
                ApplyColor(axisColor, rgb => group.AxisColor = new AxisColor { Rgb = rgb });
                ApplyColor(negativeColor, rgb => group.NegativeColor = new NegativeColor { Rgb = rgb });
                ApplyColor(markersColor, rgb => group.MarkersColor = new MarkersColor { Rgb = rgb });
                ApplyColor(highColor, rgb => group.HighMarkerColor = new HighMarkerColor { Rgb = rgb });
                ApplyColor(lowColor, rgb => group.LowMarkerColor = new LowMarkerColor { Rgb = rgb });
                ApplyColor(firstColor, rgb => group.FirstMarkerColor = new FirstMarkerColor { Rgb = rgb });
                ApplyColor(lastColor, rgb => group.LastMarkerColor = new LastMarkerColor { Rgb = rgb });

                var sparklines = new Sparklines();
                foreach (var sparklineReference in BuildSparklineReferences(dataRange, locationRange)) {
                    var sparkline = new Sparkline {
                        Formula = new OfficeFormula(sparklineReference.Formula),
                        ReferenceSequence = new OfficeReferenceSequence(sparklineReference.Location)
                    };
                    sparklines.Append(sparkline);
                }
                group.Append(sparklines);
                groups.Append(group);

                ws.Save();
            });
        }

        /// <summary>
        /// Adds line sparklines to the worksheet.
        /// </summary>
        /// <param name="dataRange">A1 range containing the data (e.g., "B2:M2").</param>
        /// <param name="locationRange">A1 range where sparklines will be placed (e.g., "N2:N2").</param>
        /// <param name="displayMarkers">Show markers for each data point.</param>
        /// <param name="displayHighLow">Show high/low markers.</param>
        /// <param name="displayFirstLast">Show first/last markers.</param>
        /// <param name="displayNegative">Show negative markers.</param>
        /// <param name="displayAxis">Show axis.</param>
        /// <param name="seriesColor">Sparkline series color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="axisColor">Axis color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="negativeColor">Negative point color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="markersColor">Markers color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="highColor">High point color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="lowColor">Low point color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="firstColor">First point color (#RRGGBB or #AARRGGBB).</param>
        /// <param name="lastColor">Last point color (#RRGGBB or #AARRGGBB).</param>
        public void AddSparklines(
            string dataRange,
            string locationRange,
            bool displayMarkers = false,
            bool displayHighLow = false,
            bool displayFirstLast = false,
            bool displayNegative = false,
            bool displayAxis = false,
            string? seriesColor = null,
            string? axisColor = null,
            string? negativeColor = null,
            string? markersColor = null,
            string? highColor = null,
            string? lowColor = null,
            string? firstColor = null,
            string? lastColor = null) {
            AddSparklines(
                dataRange,
                locationRange,
                SparklineTypeValues.Line,
                displayMarkers,
                displayHighLow,
                displayFirstLast,
                displayNegative,
                displayAxis,
                seriesColor,
                axisColor,
                negativeColor,
                markersColor,
                highColor,
                lowColor,
                firstColor,
                lastColor);
        }

        private static void ApplyColor(string? color, Action<HexBinaryValue> assign) {
            var rgb = NormalizeRgb(color);
            if (rgb != null) assign(rgb);
        }

        private static HexBinaryValue? NormalizeRgb(string? color) {
            if (string.IsNullOrWhiteSpace(color)) return null;
            var value = color!.Trim();
            if (value.StartsWith("#", StringComparison.Ordinal)) {
                value = value.Substring(1);
            }
            if (value.Length == 6) {
                value = "FF" + value;
            }
            if (value.Length != 8) {
                throw new ArgumentException($"Color '{color}' must be in #RRGGBB or #AARRGGBB format.", nameof(color));
            }
            for (int i = 0; i < value.Length; i++) {
                char c = value[i];
                bool isHexDigit = (c >= '0' && c <= '9')
                    || (c >= 'A' && c <= 'F')
                    || (c >= 'a' && c <= 'f');
                if (!isHexDigit) {
                    throw new ArgumentException(
                        $"Color '{color}' contains invalid hex character '{c}' at position {i + 1}.",
                        nameof(color));
                }
            }
            return new HexBinaryValue(value.ToUpperInvariant());
        }

        private static IReadOnlyList<SparklineReference> BuildSparklineReferences(string dataRange, string locationRange) {
            string data = dataRange.Trim();
            string location = locationRange.Trim();

            if (!TryParseSparklineRange(data, out var dataAddress) || !TryParseSparklineRange(location, out var locationAddress)) {
                return new[] { new SparklineReference(data, location) };
            }

            if (locationAddress.RowCount == 1 && locationAddress.ColumnCount == 1) {
                return new[] { new SparklineReference(data, location) };
            }

            var references = new List<SparklineReference>();
            if (locationAddress.ColumnCount == 1 && dataAddress.RowCount == locationAddress.RowCount) {
                EnsureSparklineExpansionWithinLimit(locationAddress.RowCount, nameof(locationRange));
                for (int offset = 0; offset < locationAddress.RowCount; offset++) {
                    int dataRow = dataAddress.Row1 + offset;
                    int locationRow = locationAddress.Row1 + offset;
                    references.Add(new SparklineReference(
                        dataAddress.ToReference(dataRow, dataAddress.Column1, dataRow, dataAddress.Column2),
                        locationAddress.ToReference(locationRow, locationAddress.Column1, locationRow, locationAddress.Column1)));
                }

                return references;
            }

            if (locationAddress.RowCount == 1 && dataAddress.ColumnCount == locationAddress.ColumnCount) {
                EnsureSparklineExpansionWithinLimit(locationAddress.ColumnCount, nameof(locationRange));
                for (int offset = 0; offset < locationAddress.ColumnCount; offset++) {
                    int dataColumn = dataAddress.Column1 + offset;
                    int locationColumn = locationAddress.Column1 + offset;
                    references.Add(new SparklineReference(
                        dataAddress.ToReference(dataAddress.Row1, dataColumn, dataAddress.Row2, dataColumn),
                        locationAddress.ToReference(locationAddress.Row1, locationColumn, locationAddress.Row1, locationColumn)));
                }

                return references;
            }

            throw new ArgumentException(
                "LocationRange spans multiple cells, but DataRange does not match by row or by column. " +
                "Use a single destination cell, one data row per destination row, or one data column per destination column.",
                nameof(locationRange));
        }

        private static void EnsureSparklineExpansionWithinLimit(int count, string parameterName) {
            if (count > MaximumExpandedSparklines) {
                throw new ArgumentOutOfRangeException(
                    parameterName,
                    count,
                    $"A sparkline range may expand to at most {MaximumExpandedSparklines.ToString(CultureInfo.InvariantCulture)} entries.");
            }
        }

        private static bool TryParseSparklineRange(string text, out SparklineRange range) {
            string trimmed = text.Trim();
            int separator = trimmed.LastIndexOf('!');
            string sheetPrefix = separator >= 0 ? trimmed.Substring(0, separator + 1) : string.Empty;
            string reference = separator >= 0 ? trimmed.Substring(separator + 1) : trimmed;
            reference = reference.Replace("$", string.Empty);

            if (A1.TryParseRange(reference, out int r1, out int c1, out int r2, out int c2)) {
                range = new SparklineRange(sheetPrefix, r1, c1, r2, c2);
                return true;
            }

            var cell = A1.ParseCellRef(reference);
            if (cell.Row > 0 && cell.Col > 0) {
                range = new SparklineRange(sheetPrefix, cell.Row, cell.Col, cell.Row, cell.Col);
                return true;
            }

            range = default;
            return false;
        }

        private static SparklineGroups GetOrCreateSparklineGroups(Worksheet ws) {
            const string SparklineUri = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";

            var extList = ws.Elements<DocumentFormat.OpenXml.Spreadsheet.ExtensionList>().FirstOrDefault();
            if (extList == null) {
                extList = new DocumentFormat.OpenXml.Spreadsheet.ExtensionList();
                ws.Append(extList);
            }

            var ext = extList.Elements<DocumentFormat.OpenXml.Spreadsheet.Extension>()
                .FirstOrDefault(e => string.Equals(e.Uri?.Value, SparklineUri, StringComparison.OrdinalIgnoreCase));
            if (ext == null) {
                ext = new DocumentFormat.OpenXml.Spreadsheet.Extension { Uri = SparklineUri };
                extList.Append(ext);
            }

            var groups = ext.GetFirstChild<SparklineGroups>();
            if (groups == null) {
                groups = new SparklineGroups();
                ext.Append(groups);
            }

            return groups;
        }

        private readonly struct SparklineReference {
            internal SparklineReference(string formula, string location) {
                Formula = formula;
                Location = location;
            }

            internal string Formula { get; }

            internal string Location { get; }
        }

        private readonly struct SparklineRange {
            internal SparklineRange(string sheetPrefix, int row1, int column1, int row2, int column2) {
                SheetPrefix = sheetPrefix;
                Row1 = row1;
                Column1 = column1;
                Row2 = row2;
                Column2 = column2;
            }

            internal string SheetPrefix { get; }

            internal int Row1 { get; }

            internal int Column1 { get; }

            internal int Row2 { get; }

            internal int Column2 { get; }

            internal int RowCount => Row2 - Row1 + 1;

            internal int ColumnCount => Column2 - Column1 + 1;

            internal string ToReference(int row1, int column1, int row2, int column2) {
                string start = SheetPrefix + A1.CellReference(row1, column1);
                if (row1 == row2 && column1 == column2) {
                    return start;
                }

                return start + ":" + A1.CellReference(row2, column2);
            }
        }
    }
}
