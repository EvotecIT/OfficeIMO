using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Range-based read operations for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private bool CanUseReadObjectsXmlFastPath(OfficeIMO.Excel.ExecutionMode? mode) {
            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            if (decided == OfficeIMO.Excel.ExecutionMode.Parallel) {
                return false;
            }

            return policy.OnDecision == null && CanUseXmlFastReader();
        }

        private bool TryReadObjectsDictionaryXmlStreamingFast(
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct,
            out List<Dictionary<string, object?>> result) {
            int dataRowCount = rows - 1;
            result = new List<Dictionary<string, object?>>(dataRowCount);
            var headerValues = new object?[cols];
            string[]? headers = null;
            int nextDataRow = r1 + 1;
            int lastDataRow = r1;

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;

                while (reader.Read()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                    if (rowIndex <= 0) {
                        rowIndex = nextRowIndex;
                    }

                    nextRowIndex = rowIndex + 1;
                    if (rowIndex < r1) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (rowIndex > r2) {
                        break;
                    }

                    if (rowIndex == r1) {
                        if (nextDataRow != r1 + 1 || result.Count > 0 || lastDataRow != r1) {
                            result = [];
                            return false;
                        }

                        ReadXmlRowValuesInto(reader, rowIndex, c1, c2, headerValues, ct);
                        headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                        continue;
                    }

                    if (rowIndex <= lastDataRow) {
                        result = [];
                        return false;
                    }

                    headers ??= ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                    while (nextDataRow < rowIndex) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        result.Add(CreateEmptyDictionaryRow(headers, cols));
                        nextDataRow++;
                    }

                    result.Add(ReadXmlRowIntoDictionary(reader, c1, c2, headers, cols, ct));
                    lastDataRow = rowIndex;
                    nextDataRow = rowIndex + 1;
                }

                headers ??= ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                while (nextDataRow <= r2) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    result.Add(CreateEmptyDictionaryRow(headers, cols));
                    nextDataRow++;
                }

                return result.Count == dataRowCount;
            } catch (XmlException) {
                result = [];
                return false;
            } catch (IOException) {
                result = [];
                return false;
            } catch (UnauthorizedAccessException) {
                result = [];
                return false;
            } catch (ObjectDisposedException) {
                result = [];
                return false;
            }
        }

        private Dictionary<string, object?> ReadXmlRowIntoDictionary(
            XmlReader rowReader,
            int c1,
            int c2,
            string[] headers,
            int cols,
            CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return CreateEmptyDictionaryRow(headers, cols);
            }

            if (cols > 64) {
                return ReadXmlRowIntoDictionaryBuffered(rowReader, c1, c2, headers, cols, ct);
            }

            Dictionary<string, object?>? orderedDictionary = null;
            if (TryReadXmlOrderedFullWidthRowIntoDictionary(rowReader, c1, c2, headers, cols, ct, out orderedDictionary)) {
                return orderedDictionary ?? CreateEmptyDictionaryRow(headers, cols);
            }

            return orderedDictionary ?? CreateEmptyDictionaryRow(headers, cols);
        }

        private bool TryReadXmlOrderedFullWidthRowIntoDictionary(
            XmlReader rowReader,
            int c1,
            int c2,
            string[] headers,
            int cols,
            CancellationToken ct,
            out Dictionary<string, object?>? result) {
            result = new Dictionary<string, object?>(cols, StringComparer.OrdinalIgnoreCase);
            object?[] values = new object?[cols];
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            ulong allColumnsSeen = CreateAllColumnsSeenMask(cols);
            ulong seenColumns = 0;
            int nextExpectedColumn = c1;
            bool orderedFullWidth = true;
            int visitedNodes = 0;
            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    result = orderedFullWidth && seenColumns == allColumnsSeen
                        ? result
                        : CreateDictionaryRow(headers, values, cols);
                    return true;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    orderedFullWidth = false;
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex < c1 || columnIndex > c2) {
                    orderedFullWidth = false;
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int offset = columnIndex - c1;
                if ((uint)offset >= (uint)cols) {
                    orderedFullWidth = false;
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (orderedFullWidth && columnIndex != nextExpectedColumn) {
                    orderedFullWidth = false;
                }

                object? value = ReadXmlCellValue(rowReader);
                values[offset] = value;
                if (orderedFullWidth) {
                    result.Add(headers[offset], value);
                    nextExpectedColumn++;
                }

                if (orderedFullWidth && columnIndex >= c2) {
                    SkipXmlElementContent(rowReader, depth);
                    return true;
                }

                if (!orderedFullWidth && MarkRequestedColumnSeen(offset, allColumnsSeen, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth);
                    result = CreateDictionaryRow(headers, values, cols);
                    return true;
                }
            }

            result = orderedFullWidth && seenColumns == allColumnsSeen
                ? result
                : CreateDictionaryRow(headers, values, cols);
            return true;
        }

        private Dictionary<string, object?> ReadXmlRowIntoDictionaryBuffered(
            XmlReader rowReader,
            int c1,
            int c2,
            string[] headers,
            int cols,
            CancellationToken ct) {
            object?[] values = new object?[cols];
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int visitedNodes = 0;
            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return CreateDictionaryRow(headers, values, cols);
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex < c1 || columnIndex > c2) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int offset = columnIndex - c1;
                if ((uint)offset >= (uint)cols) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                values[offset] = ReadXmlCellValue(rowReader);
            }

            return CreateDictionaryRow(headers, values, cols);
        }

        private static Dictionary<string, object?> CreateEmptyDictionaryRow(string[] headers, int columnCount) {
            var dict = new Dictionary<string, object?>(columnCount, StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < columnCount; c++) {
                dict.Add(headers[c], null);
            }

            return dict;
        }

        private static Dictionary<string, object?> CreateDictionaryRow(string[] headers, object?[]? values, int columnCount) {
            var dict = new Dictionary<string, object?>(columnCount, StringComparer.OrdinalIgnoreCase);
            if (values == null) {
                for (int c = 0; c < columnCount; c++) {
                    dict.Add(headers[c], null);
                }

                return dict;
            }

            for (int c = 0; c < columnCount; c++) {
                dict.Add(headers[c], values[c]);
            }

            return dict;
        }

        private bool TryReadObjectsDictionaryXmlFast(
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct,
            out List<Dictionary<string, object?>> result) {
            result = [];
            int dataRowCount = rows - 1;
            var headerValues = new object?[cols];
            var rowValues = dataRowCount == 0 ? Array.Empty<object?[]>() : new object?[dataRowCount][];

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                var seenRows = CreateCompletedRowTracker(rows);

                while (reader.Read()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                    if (rowIndex <= 0) {
                        rowIndex = nextRowIndex;
                    }

                    nextRowIndex = rowIndex + 1;
                    if (rowIndex < r1 || rowIndex > r2) {
                        if (rowIndex > r2 && seenRows.AllRowsSeen) {
                            break;
                        }

                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (rowIndex == r1) {
                        ReadXmlRowValuesInto(reader, rowIndex, c1, c2, headerValues, ct);
                        seenRows.MarkSeen(0);
                        continue;
                    }

                    int rowOffset = rowIndex - r1 - 1;
                    if ((uint)rowOffset >= (uint)rowValues.Length) {
                        continue;
                    }

                    object?[] values = ReadXmlRowValues(reader, rowIndex, c1, c2, cols, ct);
                    object?[]? existing = rowValues[rowOffset];
                    if (existing == null) {
                        rowValues[rowOffset] = values;
                    } else {
                        MergeRowValues(existing, values);
                    }

                    seenRows.MarkSeen(rowIndex - r1);
                }

                var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                result = new List<Dictionary<string, object?>>(dataRowCount);
                for (int r = 0; r < dataRowCount; r++) {
                    if (canCancel && (r & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    result.Add(CreateDictionaryRow(headers, rowValues[r], cols));
                }

                return true;
            } catch (XmlException) {
                result = [];
                return false;
            } catch (IOException) {
                result = [];
                return false;
            } catch (UnauthorizedAccessException) {
                result = [];
                return false;
            } catch (ObjectDisposedException) {
                result = [];
                return false;
            }

            static void MergeRowValues(object?[] target, object?[] source) {
                for (int i = 0; i < source.Length; i++) {
                    if (source[i] != null) {
                        target[i] = source[i];
                    }
                }
            }

        }

        private bool TryReadObjectsSequentialSinglePass(
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct,
            out List<Dictionary<string, object?>> result) {
            int dataRowCount = rows - 1;
            result = new List<Dictionary<string, object?>>(dataRowCount);

            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            string[]? headers = null;
            int convertedCells = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (row.RowIndex == null) {
                    return false;
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                if (rowIndex == r1) {
                    var headerValues = new object?[cols];
                    foreach (var cell in row.Elements<Cell>()) {
                        if (canCancel && (++convertedCells & 1023) == 0) {
                            ct.ThrowIfCancellationRequested();
                        }

                        int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                        if (columnIndex < c1 || columnIndex > c2) {
                            continue;
                        }

                        int cc = columnIndex - c1;
                        if ((uint)cc >= (uint)cols) {
                            continue;
                        }

                        if (TryConvertCell(cell, out object? value)) {
                            headerValues[cc] = value;
                        }
                    }

                    headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                    for (int r = 0; r < dataRowCount; r++) {
                        if (canCancel && (r & 1023) == 0) {
                            ct.ThrowIfCancellationRequested();
                        }

                        result.Add(CreateEmptyRow(headers));
                    }

                    continue;
                }

                if (headers == null) {
                    return false;
                }

                int rr = rowIndex - r1 - 1;
                if ((uint)rr >= (uint)result.Count) {
                    continue;
                }

                var dict = result[rr];
                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancel && (++convertedCells & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (columnIndex < c1 || columnIndex > c2) {
                        continue;
                    }

                    int cc = columnIndex - c1;
                    if ((uint)cc >= (uint)cols) {
                        continue;
                    }

                    if (TryConvertCell(cell, out object? value)) {
                        dict[headers[cc]] = value;
                    }
                }
            }

            return headers != null;

            Dictionary<string, object?> CreateEmptyRow(string[] rowHeaders) {
                var dict = new Dictionary<string, object?>(cols, StringComparer.OrdinalIgnoreCase);
                for (int c = 0; c < cols; c++) {
                    dict.Add(rowHeaders[c], null);
                }

                return dict;
            }
        }

    }
}
