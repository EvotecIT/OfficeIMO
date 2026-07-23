using System.IO;
using System.Globalization;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Used-range read operations for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Reads the worksheet used range into a dense two-dimensional array of typed values.
        /// Table-backed worksheets are discovered and materialized in one worksheet pass.
        /// </summary>
        /// <param name="ct">Cancellation token observed while reading the worksheet.</param>
        /// <returns>Typed matrix populated from the worksheet used range.</returns>
        public object?[,] ReadUsedRange(CancellationToken ct = default) {
            if (CanUseXmlFastReader()
                && CanUseAutomaticXmlReadFastPath(_opt.Execution)
                && TryReadTableBackedUsedRangeXmlFast(ct, out object?[,] values)) {
                return values;
            }

            string usedRange = GetUsedRangeA1();
            return ReadRange(usedRange, ct: ct);
        }

        private bool TryReadTableBackedUsedRangeXmlFast(CancellationToken ct, out object?[,] values) {
            values = new object?[0, 0];
            if (!TryGetWorksheetDimensionReferenceFromXml(out string dimensionReference)
                || !TryGetTableBackedDimensionReference(dimensionReference, out string tableBackedReference)
                || !A1.TryParseRange(tableBackedReference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return false;
            }

            int height = lastRow - firstRow + 1;
            int width = lastColumn - firstColumn + 1;
            if (height <= 0 || width <= 0) {
                return false;
            }

            if (_opt.MaxRangeCells <= 0) {
                throw new ArgumentOutOfRangeException(nameof(_opt.MaxRangeCells), "Maximum dense range cell count must be positive.");
            }

            long cellCount = (long)height * width;
            if (cellCount > _opt.MaxRangeCells) {
                throw new InvalidDataException(
                    $"Range '{tableBackedReference}' contains {cellCount.ToString(CultureInfo.InvariantCulture)} cells, exceeding the configured limit of {_opt.MaxRangeCells.ToString(CultureInfo.InvariantCulture)}.");
            }

            var result = new object?[height, width];
            if (TryFillRangeUtf8Fast(result, firstRow, firstColumn, lastRow, lastColumn, ct, requireAllWorksheetCellsWithinRange: true)) {
                _usedRangeA1 = tableBackedReference;
                values = result;
                return true;
            }

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!TryPrepareWorksheetStream(stream)) {
                    return false;
                }

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
                    bool hasExplicitRowIndex = rowIndex > 0;
                    if (!hasExplicitRowIndex) {
                        rowIndex = nextRowIndex;
                    }

                    nextRowIndex = rowIndex + 1;
                    if (reader.IsEmptyElement) {
                        continue;
                    }

                    int rowDepth = reader.Depth;
                    int nextColumnIndex = 1;
                    while (reader.Read()) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        if (reader.NodeType == XmlNodeType.EndElement
                            && reader.Depth == rowDepth
                            && reader.LocalName == "row") {
                            break;
                        }

                        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c") {
                            continue;
                        }

                        int cellRow = rowIndex;
                        int cellColumn;
                        if (hasExplicitRowIndex) {
                            cellColumn = GetXmlCellColumnIndex(reader, ref nextColumnIndex);
                        } else if (A1.TryParseCellReferenceFast(reader.GetAttribute("r"), out int parsedRow, out int parsedColumn)) {
                            if (parsedRow > 0) {
                                cellRow = parsedRow;
                            }

                            cellColumn = parsedColumn;
                            nextColumnIndex = parsedColumn + 1;
                        } else {
                            cellColumn = nextColumnIndex++;
                        }

                        if (cellColumn <= 0
                            || cellRow < firstRow
                            || cellRow > lastRow
                            || cellColumn < firstColumn
                            || cellColumn > lastColumn) {
                            return false;
                        }

                        result[cellRow - firstRow, cellColumn - firstColumn] = ReadXmlCellValue(reader, reader.GetAttribute("t"));
                    }
                }

                _usedRangeA1 = tableBackedReference;
                values = result;
                return true;
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }
        }
    }
}
