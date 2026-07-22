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
        private const int DenseSnapshotCapacityLimit = 100_000;
        // DataTable materialization already carries row storage cost; keep the single-pass XML buffer active
        // through larger normal sheets to avoid a slower second worksheet scan.
        private const int DataTableBufferedSinglePassCapacityLimit = 3_000_000;
        private const int SparseReadInitialBufferCapacity = 64;
        private const int XmlFastCompletedRowTrackingLimit = 4096;

        /// <summary>
        /// Returns the used range of the worksheet as an A1 string (e.g., "A1:C10").
        /// If the sheet is empty, returns "A1:A1".
        /// </summary>
        public string GetUsedRangeA1() {
            if (_usedRangeA1 != null) {
                return _usedRangeA1;
            }

            if (_canStreamWorksheetPart
                && TryGetWorksheetDimensionReferenceFromXml(out string dimensionReference)
                && TryGetTableBackedDimensionReference(dimensionReference, out string tableBackedReference)
                && TryWorksheetCellsFitWithinRangeFromXml(tableBackedReference)) {
                _usedRangeA1 = tableBackedReference;
                return tableBackedReference;
            }

            if (_canStreamWorksheetPart
                && TryComputeUsedRangeReferenceFromXml(out string usedRangeReference)) {
                _usedRangeA1 = usedRangeReference;
                return usedRangeReference;
            }

            string reference = ExcelSheet.ComputeSheetDimensionReference(WorksheetRoot);
            string usedRange = reference.IndexOf(":", StringComparison.Ordinal) >= 0 ? reference : reference + ":" + reference;
            _usedRangeA1 = usedRange;
            return usedRange;
        }

        private bool TryGetWorksheetDimensionReferenceFromXml(out string reference) {
            reference = string.Empty;
            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!TryPrepareWorksheetStream(stream)) {
                    return false;
                }

                using var reader = OpenWorksheetXmlReader(stream);
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element) {
                        continue;
                    }

                    if (reader.LocalName == "dimension") {
                        return TryNormalizeWorksheetDimensionReference(reader.GetAttribute("ref"), out reference);
                    }

                    if (reader.LocalName == "sheetData") {
                        return false;
                    }
                }
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }

            return false;
        }

        private bool TryComputeUsedRangeReferenceFromXml(out string reference) {
            reference = string.Empty;
            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!TryPrepareWorksheetStream(stream)) {
                    return false;
                }

                using var reader = OpenWorksheetXmlReader(stream);

                int minRow = int.MaxValue;
                int minColumn = int.MaxValue;
                int maxRow = 0;
                int maxColumn = 0;
                int nextRowIndex = 1;

                while (reader.Read()) {
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

                    int rowMinRow = hasExplicitRowIndex ? rowIndex : int.MaxValue;
                    int rowMaxRow = hasExplicitRowIndex ? rowIndex : 0;
                    int rowMinColumn = int.MaxValue;
                    int rowMaxColumn = 0;
                    int rowDepth = reader.Depth;
                    int nextColumnIndex = 1;
                    bool advanceReader = true;
                    while (advanceReader ? reader.Read() : !reader.EOF) {
                        advanceReader = true;
                        if (reader.NodeType == XmlNodeType.EndElement
                            && reader.Depth == rowDepth
                            && reader.LocalName == "row") {
                            break;
                        }

                        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c") {
                            continue;
                        }

                        int column = 0;
                        if (hasExplicitRowIndex) {
                            column = GetXmlCellColumnIndex(reader, ref nextColumnIndex);
                        } else if (A1.TryParseCellReferenceFast(reader.GetAttribute("r"), out int parsedRow, out int parsedColumn)) {
                            column = parsedColumn;
                            if (parsedRow > 0) {
                                if (parsedRow < rowMinRow) rowMinRow = parsedRow;
                                if (parsedRow > rowMaxRow) rowMaxRow = parsedRow;
                            }

                            nextColumnIndex = parsedColumn + 1;
                        }

                        if (column <= 0) {
                            column = nextColumnIndex;
                            nextColumnIndex = column + 1;
                        }

                        if (column > 0) {
                            if (column < rowMinColumn) rowMinColumn = column;
                            if (column > rowMaxColumn) rowMaxColumn = column;
                        }

                        if (!reader.IsEmptyElement) {
                            reader.Skip();
                            advanceReader = false;
                        }
                    }

                    if (rowMaxColumn <= 0) {
                        continue;
                    }

                    if (rowMaxRow <= 0) {
                        rowMinRow = rowIndex;
                        rowMaxRow = rowIndex;
                    }

                    if (rowMinRow < minRow) minRow = rowMinRow;
                    if (rowMaxRow > maxRow) maxRow = rowMaxRow;
                    if (rowMinColumn < minColumn) minColumn = rowMinColumn;
                    if (rowMaxColumn > maxColumn) maxColumn = rowMaxColumn;
                    if (!hasExplicitRowIndex) {
                        nextRowIndex = rowMaxRow + 1;
                    }
                }

                if (maxRow <= 0 || maxColumn <= 0) {
                    return false;
                }

                reference = A1.CellReference(minRow, minColumn) + ":" + A1.CellReference(maxRow, maxColumn);
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

        private bool TryWorksheetCellsFitWithinRangeFromXml(string reference) {
            if (!A1.TryParseRange(reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return false;
            }

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!TryPrepareWorksheetStream(stream)) {
                    return false;
                }

                using var reader = OpenWorksheetXmlReader(stream);
                int nextRowIndex = 1;
                while (reader.Read()) {
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
                    bool advanceReader = true;
                    while (advanceReader ? reader.Read() : !reader.EOF) {
                        advanceReader = true;
                        if (reader.NodeType == XmlNodeType.EndElement
                            && reader.Depth == rowDepth
                            && reader.LocalName == "row") {
                            break;
                        }

                        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c") {
                            continue;
                        }

                        int cellRow = rowIndex;
                        int cellColumn = 0;
                        if (hasExplicitRowIndex) {
                            cellColumn = GetXmlCellColumnIndex(reader, ref nextColumnIndex);
                        } else if (A1.TryParseCellReferenceFast(reader.GetAttribute("r"), out int parsedRow, out int parsedColumn)) {
                            if (parsedRow > 0) {
                                cellRow = parsedRow;
                            }

                            cellColumn = parsedColumn;
                            nextColumnIndex = parsedColumn + 1;
                        }

                        if (cellColumn <= 0) {
                            cellColumn = nextColumnIndex;
                            nextColumnIndex = cellColumn + 1;
                        }

                        if (cellRow < firstRow || cellRow > lastRow || cellColumn < firstColumn || cellColumn > lastColumn) {
                            return false;
                        }

                        if (!reader.IsEmptyElement) {
                            reader.Skip();
                            advanceReader = false;
                        }
                    }
                }

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

        private static bool TryGetWorksheetDimensionReference(Worksheet worksheet, out string reference) {
            reference = string.Empty;
            string? rawReference = worksheet.SheetDimension?.Reference?.Value;
            return TryNormalizeWorksheetDimensionReference(rawReference, out reference);
        }

        private bool TryGetTableBackedDimensionReference(string dimensionReference, out string reference) {
            reference = string.Empty;
            if (!A1.TryParseRange(dimensionReference, out int dimensionFirstRow, out int dimensionFirstColumn, out int dimensionLastRow, out int dimensionLastColumn)) {
                return false;
            }

            int minRow = int.MaxValue;
            int minColumn = int.MaxValue;
            int maxRow = 0;
            int maxColumn = 0;

            try {
                foreach (var tablePart in _wsPart.TableDefinitionParts) {
                    string? tableReference = TryGetTableReferenceXmlFast(tablePart, out string xmlTableReference)
                        ? xmlTableReference
                        : tablePart.Table?.Reference?.Value;
                    if (string.IsNullOrWhiteSpace(tableReference)
                        || !A1.TryParseRange(tableReference!, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                        return false;
                    }

                    if (firstRow < minRow) minRow = firstRow;
                    if (firstColumn < minColumn) minColumn = firstColumn;
                    if (lastRow > maxRow) maxRow = lastRow;
                    if (lastColumn > maxColumn) maxColumn = lastColumn;
                }
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }

            if (maxRow <= 0
                || minRow != dimensionFirstRow
                || minColumn != dimensionFirstColumn
                || maxRow != dimensionLastRow
                || maxColumn != dimensionLastColumn) {
                return false;
            }

            reference = dimensionReference;
            return true;
        }

        private static bool TryGetTableReferenceXmlFast(TableDefinitionPart tablePart, out string reference) {
            reference = string.Empty;
            try {
                using var stream = tablePart.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, WorksheetXmlReaderSettings);
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "table") {
                        continue;
                    }

                    string? tableReference = reader.GetAttribute("ref");
                    if (string.IsNullOrWhiteSpace(tableReference)) {
                        return false;
                    }

                    reference = tableReference!;
                    return true;
                }
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }

            return false;
        }

        private static bool TryNormalizeWorksheetDimensionReference(string? rawReference, out string reference) {
            reference = string.Empty;
            if (string.IsNullOrWhiteSpace(rawReference)) {
                return false;
            }

            rawReference = rawReference!.Trim();
            if (rawReference.Equals("A1", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (rawReference.IndexOf(':') >= 0) {
                if (!A1.TryParseRange(rawReference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
                    || firstRow <= 0
                    || firstColumn <= 0
                    || lastRow < firstRow
                    || lastColumn < firstColumn) {
                    return false;
                }

                reference = rawReference;
                return true;
            }

            if (!A1.TryParseCellReferenceFast(rawReference, out int row, out int column)
                || row <= 0
                || column <= 0) {
                return false;
            }

            reference = rawReference + ":" + rawReference;
            return true;
        }
        /// <summary>
        /// Reads a rectangular A1 range (e.g., "A1:C10") into a dense 2D array of typed values.
        /// </summary>
        public object?[,] ReadRange(string a1Range, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            var height = r2 - r1 + 1;
            var width = c2 - c1 + 1;
            long cellCount = (long)height * width;
            if (_opt.MaxRangeCells <= 0) {
                throw new ArgumentOutOfRangeException(nameof(_opt.MaxRangeCells), "Maximum dense range cell count must be positive.");
            }

            if (cellCount > _opt.MaxRangeCells) {
                throw new InvalidDataException(
                    $"Range '{a1Range}' contains {cellCount.ToString(CultureInfo.InvariantCulture)} cells, exceeding the configured limit of {_opt.MaxRangeCells.ToString(CultureInfo.InvariantCulture)}.");
            }

            var result = new object?[height, width];

            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            int workload = checked((int)cellCount);
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) {
                if (CanUseAutomaticXmlReadFastPath(policy)) {
                    if ((ShouldAttemptUtf8Range(r1, r2) && RangeReachesDeclaredWorksheetEnd(r2) && TryFillRangeUtf8Fast(result, r1, c1, r2, c2, ct))
                        || TryFillRangeXmlFast(result, r1, c1, r2, c2, ct)) {
                        return result;
                    }
                }

                decided = policy.Decide("ReadRange", workload);
            }

            if (decided == OfficeIMO.Excel.ExecutionMode.Sequential) {
                if ((ShouldAttemptUtf8Range(r1, r2) && RangeReachesDeclaredWorksheetEnd(r2) && TryFillRangeUtf8Fast(result, r1, c1, r2, c2, ct))
                    || TryFillRangeXmlFast(result, r1, c1, r2, c2, ct)) {
                    return result;
                }

                FillRangeSequential(result, r1, c1, r2, c2, ct);
                return result;
            }

            var raw = SnapshotAndConvertRangeCells(r1, c1, r2, c2, "ReadRange", decided, ct, workload);

            foreach (var cell in raw) {
                var rr = cell.Row - r1;
                var cc = cell.Col - c1;
                if ((uint)rr < (uint)height && (uint)cc < (uint)width)
                    result[rr, cc] = cell.TypedValue;
            }

            return result;
        }
    }
}
