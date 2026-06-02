using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool CanSkipStableAutoFitColumns(IReadOnlyList<int>? requestedColumns) {
            if (_hasWorksheetMutations || _excelDocument.IsPackageDirty) {
                return false;
            }

            if (TryCanSkipStableAutoFitColumnsFromWorksheetXml(requestedColumns, out bool canSkip)) {
                return canSkip;
            }

            var worksheet = WorksheetRoot;
            var columns = worksheet.GetFirstChild<Columns>();
            if (columns == null) {
                return false;
            }

            IReadOnlyList<int> targetColumns;
            if (requestedColumns != null) {
                targetColumns = requestedColumns;
            } else if (TryGetDimensionColumnBounds(worksheet, out int firstColumn, out int lastColumn)) {
                targetColumns = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1).ToArray();
            } else if (TryGetSheetDataColumnBounds(worksheet, out firstColumn, out lastColumn)) {
                targetColumns = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1).ToArray();
            } else {
                return false;
            }

            bool[] stableColumns = new bool[A1.MaxColumns + 1];
            foreach (var column in columns.Elements<Column>()) {
                if (column.Width == null
                    || column.CustomWidth?.Value != true
                    || column.BestFit?.Value != true) {
                    continue;
                }

                uint min = column.Min?.Value ?? 0U;
                uint max = column.Max?.Value ?? 0U;
                if (min == 0U || max < min || min > A1.MaxColumns) {
                    continue;
                }

                int start = (int)Math.Max(1U, min);
                int end = (int)Math.Min((uint)A1.MaxColumns, max);
                for (int i = start; i <= end; i++) {
                    stableColumns[i] = true;
                }
            }

            foreach (int columnIndex in targetColumns) {
                if (columnIndex <= 0
                    || columnIndex > A1.MaxColumns
                    || !stableColumns[columnIndex]) {
                    return false;
                }
            }

            return true;
        }

        private bool TryCanSkipStableAutoFitColumnsFromWorksheetXml(IReadOnlyList<int>? requestedColumns, out bool canSkip) {
            canSkip = false;

            try {
                using var stream = _worksheetPart.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                    IgnoreWhitespace = true
                });

                int firstColumn = 0;
                int lastColumn = 0;
                bool hasDimensionBounds = false;
                bool sawColumns = false;
                var stableSpans = new List<(int Start, int End)>();

                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element) {
                        continue;
                    }

                    if (reader.LocalName == "dimension") {
                        string? reference = reader.GetAttribute("ref");
                        if (TryParseDimensionColumnBounds(reference, out firstColumn, out lastColumn)) {
                            hasDimensionBounds = true;
                        }
                        continue;
                    }

                    if (reader.LocalName == "cols") {
                        sawColumns = true;
                        if (reader.IsEmptyElement) {
                            continue;
                        }

                        int depth = reader.Depth;
                        while (reader.Read()) {
                            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == "cols") {
                                break;
                            }

                            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "col") {
                                continue;
                            }

                            if (!TryReadStableColumnSpan(reader, out int start, out int end)) {
                                continue;
                            }

                            stableSpans.Add((start, end));
                        }

                        continue;
                    }

                    if (reader.LocalName == "sheetData") {
                        break;
                    }
                }

                if (!sawColumns) {
                    canSkip = false;
                    return true;
                }

                if (requestedColumns != null) {
                    canSkip = AreRequestedColumnsCovered(requestedColumns, stableSpans);
                    return true;
                }

                if (!hasDimensionBounds) {
                    return false;
                }

                canSkip = AreColumnRangeCovered(firstColumn, lastColumn, stableSpans);
                return true;
            } catch (IOException) {
                return false;
            } catch (XmlException) {
                return false;
            } catch (InvalidOperationException) {
                return false;
            }
        }

        private static bool TryParseDimensionColumnBounds(string? reference, out int firstColumn, out int lastColumn) {
            firstColumn = 0;
            lastColumn = 0;

            if (string.IsNullOrWhiteSpace(reference)) {
                return false;
            }

            if (reference!.IndexOf(':') >= 0) {
                return A1.TryParseRange(reference, out _, out firstColumn, out _, out lastColumn)
                    && firstColumn > 0
                    && lastColumn >= firstColumn;
            }

            var parsed = A1.ParseCellRef(reference);
            firstColumn = parsed.Col;
            lastColumn = parsed.Col;
            return firstColumn > 0;
        }

        private static bool TryReadStableColumnSpan(XmlReader reader, out int start, out int end) {
            start = 0;
            end = 0;

            string? minText = reader.GetAttribute("min");
            string? maxText = reader.GetAttribute("max");
            if (!uint.TryParse(minText, NumberStyles.None, CultureInfo.InvariantCulture, out uint min)
                || !uint.TryParse(maxText, NumberStyles.None, CultureInfo.InvariantCulture, out uint max)
                || min == 0U
                || max < min
                || min > A1.MaxColumns) {
                return false;
            }

            if (reader.GetAttribute("width") == null
                || !IsOpenXmlBooleanTrue(reader.GetAttribute("customWidth"))
                || !IsOpenXmlBooleanTrue(reader.GetAttribute("bestFit"))) {
                return false;
            }

            start = (int)Math.Max(1U, min);
            end = (int)Math.Min((uint)A1.MaxColumns, max);
            return end >= start;
        }

        private static bool IsOpenXmlBooleanTrue(string? value)
            => string.Equals(value, "1", StringComparison.Ordinal)
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);

        private static bool AreRequestedColumnsCovered(IReadOnlyList<int> requestedColumns, List<(int Start, int End)> stableSpans) {
            foreach (int columnIndex in requestedColumns) {
                if (columnIndex <= 0
                    || columnIndex > A1.MaxColumns
                    || !IsColumnCovered(columnIndex, stableSpans)) {
                    return false;
                }
            }

            return true;
        }

        private static bool AreColumnRangeCovered(int firstColumn, int lastColumn, List<(int Start, int End)> stableSpans) {
            if (firstColumn <= 0 || lastColumn < firstColumn || lastColumn > A1.MaxColumns) {
                return false;
            }

            for (int columnIndex = firstColumn; columnIndex <= lastColumn; columnIndex++) {
                if (!IsColumnCovered(columnIndex, stableSpans)) {
                    return false;
                }
            }

            return true;
        }

        private static bool IsColumnCovered(int columnIndex, List<(int Start, int End)> stableSpans) {
            for (int i = 0; i < stableSpans.Count; i++) {
                var span = stableSpans[i];
                if (columnIndex >= span.Start && columnIndex <= span.End) {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetDimensionColumnBounds(Worksheet worksheet, out int firstColumn, out int lastColumn) {
            firstColumn = 0;
            lastColumn = 0;
            string? reference = worksheet.SheetDimension?.Reference?.Value;
            if (string.IsNullOrWhiteSpace(reference)) {
                return false;
            }

            if (reference!.IndexOf(':') >= 0) {
                if (!A1.TryParseRange(reference, out _, out firstColumn, out _, out lastColumn)) {
                    return false;
                }
            } else {
                var parsed = A1.ParseCellRef(reference);
                firstColumn = parsed.Col;
                lastColumn = parsed.Col;
            }

            return firstColumn > 0 && lastColumn >= firstColumn;
        }

        private static bool TryGetSheetDataColumnBounds(Worksheet worksheet, out int firstColumn, out int lastColumn) {
            firstColumn = int.MaxValue;
            lastColumn = 0;

            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                firstColumn = 0;
                return false;
            }

            foreach (var row in sheetData.Elements<Row>()) {
                foreach (var cell in row.Elements<Cell>()) {
                    string? reference = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(reference)) {
                        continue;
                    }

                    var parsed = A1.ParseCellRef(reference!);
                    if (parsed.Col <= 0) {
                        continue;
                    }

                    if (parsed.Col < firstColumn) {
                        firstColumn = parsed.Col;
                    }

                    if (parsed.Col > lastColumn) {
                        lastColumn = parsed.Col;
                    }
                }
            }

            if (firstColumn == int.MaxValue) {
                firstColumn = 0;
            }

            return firstColumn > 0 && lastColumn >= firstColumn;
        }
    }
}
