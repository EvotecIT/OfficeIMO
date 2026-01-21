using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;
using OfficeReferenceSequence = DocumentFormat.OpenXml.Office.Excel.ReferenceSequence;
using System.Globalization;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Helpers for adding sparklines to worksheets.
    /// </summary>
    public partial class ExcelSheet {
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
                var ws = _worksheetPart.Worksheet;
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
                var sparkline = new Sparkline {
                    Formula = new OfficeFormula(dataRange.Trim()),
                    ReferenceSequence = new OfficeReferenceSequence(locationRange.Trim())
                };
                sparklines.Append(sparkline);
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
            return new HexBinaryValue(value.ToUpperInvariant());
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
    }
}
