using OfficeIMO.OpenXml.Internal;
using System.Globalization;

namespace OfficeIMO.Excel {
    // Helpers for adding sparklines to worksheets.
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
        /// <param name="seriesColor">Optional literal color or theme role.</param>
        /// <param name="axisColor">Optional literal color or theme role.</param>
        /// <param name="negativeColor">Optional literal color or theme role.</param>
        /// <param name="markersColor">Optional literal color or theme role.</param>
        /// <param name="highColor">Optional literal color or theme role.</param>
        /// <param name="lowColor">Optional literal color or theme role.</param>
        /// <param name="firstColor">Optional literal color or theme role.</param>
        /// <param name="lastColor">Optional literal color or theme role.</param>
        public ExcelSparklineGroup AddSparklines(
            string dataRange,
            string locationRange,
            ExcelSparklineType type,
            bool displayMarkers = false,
            bool displayHighLow = false,
            bool displayFirstLast = false,
            bool displayNegative = false,
            bool displayAxis = false,
            OfficeColor? seriesColor = null,
            OfficeColor? axisColor = null,
            OfficeColor? negativeColor = null,
            OfficeColor? markersColor = null,
            OfficeColor? highColor = null,
            OfficeColor? lowColor = null,
            OfficeColor? firstColor = null,
            OfficeColor? lastColor = null) {
            ArgumentException.ThrowIfNullOrWhiteSpace(dataRange);
            ArgumentException.ThrowIfNullOrWhiteSpace(locationRange);

            string typeName = SparklineTypeName(type);
            IReadOnlyList<SparklineReference> references = BuildSparklineReferences(dataRange, locationRange);
            ExcelSparklineGroup? created = null;
            EditWorksheetObject(worksheet => {
                OfficeXmlNode groups = GetOrCreateSparklineGroups(worksheet);
                OfficeXmlNode group = AppendSparklineElement(groups, "sparklineGroup");
                group.SetAttribute("type", typeName);
                if (displayMarkers) group.SetBooleanAttribute("markers", true);
                if (displayHighLow) {
                    group.SetBooleanAttribute("high", true);
                    group.SetBooleanAttribute("low", true);
                }
                if (displayFirstLast) {
                    group.SetBooleanAttribute("first", true);
                    group.SetBooleanAttribute("last", true);
                }
                if (displayNegative) group.SetBooleanAttribute("negative", true);
                if (displayAxis) group.SetBooleanAttribute("displayXAxis", true);

                ApplySparklineColor(group, "colorSeries", seriesColor);
                ApplySparklineColor(group, "colorNegative", negativeColor);
                ApplySparklineColor(group, "colorAxis", axisColor);
                ApplySparklineColor(group, "colorMarkers", markersColor);
                ApplySparklineColor(group, "colorFirst", firstColor);
                ApplySparklineColor(group, "colorLast", lastColor);
                ApplySparklineColor(group, "colorHigh", highColor);
                ApplySparklineColor(group, "colorLow", lowColor);

                OfficeXmlNode sparklines = AppendSparklineElement(group, "sparklines");
                foreach (SparklineReference reference in references) {
                    OfficeXmlNode sparkline = AppendSparklineElement(sparklines, "sparkline");
                    AppendSparklineText(sparkline, "f", reference.Formula);
                    AppendSparklineText(sparkline, "sqref", reference.Location);
                }
                created = new ExcelSparklineGroup(this, group, references[0].Formula, references[0].Location, type);
                return true;
            }, save: true);
            return created!;
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
        /// <param name="seriesColor">Optional literal color or theme role.</param>
        /// <param name="axisColor">Optional literal color or theme role.</param>
        /// <param name="negativeColor">Optional literal color or theme role.</param>
        /// <param name="markersColor">Optional literal color or theme role.</param>
        /// <param name="highColor">Optional literal color or theme role.</param>
        /// <param name="lowColor">Optional literal color or theme role.</param>
        /// <param name="firstColor">Optional literal color or theme role.</param>
        /// <param name="lastColor">Optional literal color or theme role.</param>
        public ExcelSparklineGroup AddSparklines(
            string dataRange,
            string locationRange,
            bool displayMarkers = false,
            bool displayHighLow = false,
            bool displayFirstLast = false,
            bool displayNegative = false,
            bool displayAxis = false,
            OfficeColor? seriesColor = null,
            OfficeColor? axisColor = null,
            OfficeColor? negativeColor = null,
            OfficeColor? markersColor = null,
            OfficeColor? highColor = null,
            OfficeColor? lowColor = null,
            OfficeColor? firstColor = null,
            OfficeColor? lastColor = null) {
            return AddSparklines(
                dataRange,
                locationRange,
                ExcelSparklineType.Line,
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

        private static void ApplySparklineColor(OfficeXmlNode group, string name, OfficeColor? color) {
            if (color.HasValue) AppendSparklineElement(group, name).SetAttribute("rgb", color.Value.Resolve(OfficeTheme.Default).ToArgbHex());
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

            if (ExcelA1.TryParseRange(reference, out int r1, out int c1, out int r2, out int c2)) {
                range = new SparklineRange(sheetPrefix, r1, c1, r2, c2);
                return true;
            }

            var cell = ExcelA1.ParseCellRef(reference);
            if (cell.Row > 0 && cell.Col > 0) {
                range = new SparklineRange(sheetPrefix, cell.Row, cell.Col, cell.Row, cell.Col);
                return true;
            }

            range = default;
            return false;
        }

        private const string SparklineExtensionUri = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
        private const string SparklineNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        private const string ExcelFormulaNamespace = "http://schemas.microsoft.com/office/excel/2006/main";

        private static OfficeXmlNode AppendSparklineElement(OfficeXmlNode parent, string name) =>
            parent.Append(parent.CreateElement(SparklineNamespace, name, "x14"));

        private static void AppendSparklineText(OfficeXmlNode parent, string name, string text) =>
            parent.Append(parent.CreateElement(ExcelFormulaNamespace, name, "xm")).SetInnerText(text);

        private static OfficeXmlNode FindSparklineExtension(OfficeXmlNode extensionList) {
            foreach (OfficeXmlNode extension in extensionList.Children(WorksheetMarkup, "ext")) {
                if (string.Equals(extension.Attribute("uri"), SparklineExtensionUri, StringComparison.OrdinalIgnoreCase)) return extension;
            }
            return default;
        }

        private static OfficeXmlNode GetOrCreateSparklineGroups(OfficeXmlNode worksheet) {
            OfficeXmlNode list = AddWorksheetObject(worksheet, "extLst");
            OfficeXmlNode extension = FindSparklineExtension(list);
            if (extension.IsNone) {
                extension = list.Append(list.CreateElement(WorksheetMarkup, "ext"));
                extension.SetAttribute("uri", SparklineExtensionUri);
            }
            OfficeXmlNode groups = extension.Child(SparklineNamespace, "sparklineGroups");
            return groups.IsNone ? AppendSparklineElement(extension, "sparklineGroups") : groups;
        }

        private static IEnumerable<OfficeXmlNode> FindSparklineGroups(OfficeXmlNode worksheet) {
            OfficeXmlNode list = worksheet.Child(WorksheetMarkup, "extLst");
            foreach (OfficeXmlNode extension in list.Children(WorksheetMarkup, "ext")) {
                if (!string.Equals(extension.Attribute("uri"), SparklineExtensionUri, StringComparison.OrdinalIgnoreCase)) continue;
                OfficeXmlNode groups = extension.Child(SparklineNamespace, "sparklineGroups");
                foreach (OfficeXmlNode group in groups.Children()) {
                    if (group.Is(SparklineNamespace, "sparklineGroup")) yield return group;
                }
            }
        }

        private static IEnumerable<OfficeXmlNode> FindSparklines(OfficeXmlNode group) {
            foreach (OfficeXmlNode sparkline in group.Child(SparklineNamespace, "sparklines").Children()) {
                if (sparkline.Is(SparklineNamespace, "sparkline")) yield return sparkline;
            }
        }

        internal static string SparklineTypeName(ExcelSparklineType type) => type switch {
            ExcelSparklineType.Line => "line",
            ExcelSparklineType.Column => "column",
            ExcelSparklineType.Stacked => "stacked",
            _ => throw new ArgumentOutOfRangeException(nameof(type))
        };

        internal static ExcelSparklineType ReadSparklineType(OfficeXmlNode group) => group.Attribute("type") switch {
            "column" => ExcelSparklineType.Column,
            "stacked" => ExcelSparklineType.Stacked,
            _ => ExcelSparklineType.Line
        };

        internal ExcelSparklineType ReadSparklineGroupType(string formula, string location, ExcelSparklineType previous) =>
            ReadWorksheetObject(worksheet => {
                OfficeXmlNode group = FindSparklineByReference(worksheet, formula, location);
                return group.IsNone ? previous : ReadSparklineType(group);
            });

        internal void EditSparklineGroupType(string formula, string location, string type) =>
            EditWorksheetObject(worksheet => {
                OfficeXmlNode group = FindSparklineByReference(worksheet, formula, location);
                if (group.IsNone) return false;
                group.SetAttribute("type", type);
                return true;
            });

        private static OfficeXmlNode FindSparklineByReference(OfficeXmlNode worksheet, string formula, string location) {
            OfficeXmlNode match = default;
            foreach (OfficeXmlNode group in FindSparklineGroups(worksheet)) {
                foreach (OfficeXmlNode sparkline in FindSparklines(group)) {
                    if (sparkline.Child(ExcelFormulaNamespace, "f").InnerText == formula &&
                        sparkline.Child(ExcelFormulaNamespace, "sqref").InnerText == location) match = group;
                }
            }
            return match;
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
                string start = SheetPrefix + ExcelA1.CellReference(row1, column1);
                if (row1 == row2 && column1 == column2) {
                    return start;
                }

                return start + ":" + ExcelA1.CellReference(row2, column2);
            }
        }
    }
}
