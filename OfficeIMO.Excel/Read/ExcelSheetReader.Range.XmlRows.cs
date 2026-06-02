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
        private bool CanUseXmlFastReader() {
            return _opt.CellValueConverter == null
                && _opt.Culture == CultureInfo.InvariantCulture
                && CanStreamWorksheetPart();
        }

        private bool CanAttemptXmlFastReader() {
            return _opt.CellValueConverter == null
                && _opt.Culture == CultureInfo.InvariantCulture
                && _canStreamWorksheetPart;
        }

        private bool CanUseAutomaticXmlReadFastPath(ExecutionPolicy policy) {
            return policy.OnDecision == null;
        }

        private bool CanUseDataTableXmlBufferedReader() {
            return (_opt.CellValueConverter != null || _opt.Culture == CultureInfo.InvariantCulture)
                && CanStreamWorksheetPart();
        }

        private static CompletedRowTracker CreateCompletedRowTracker(int rowCount) {
            return new CompletedRowTracker(rowCount);
        }

        private struct CompletedRowTracker {
            private readonly int _rowCount;
            private bool[]? _seenRows;
            private ulong _seenRowMask0;
            private ulong _seenRowMask1;
            private ulong _seenRowMask2;
            private ulong _seenRowMask3;
            private int _seenRowCount;

            internal CompletedRowTracker(int rowCount) {
                if (rowCount <= 0 || rowCount > XmlFastCompletedRowTrackingLimit) {
                    _rowCount = 0;
                    _seenRows = null;
                    _seenRowMask0 = 0;
                    _seenRowMask1 = 0;
                    _seenRowMask2 = 0;
                    _seenRowMask3 = 0;
                    _seenRowCount = 0;
                    return;
                }

                _rowCount = rowCount;
                _seenRows = null;
                _seenRowMask0 = 0;
                _seenRowMask1 = 0;
                _seenRowMask2 = 0;
                _seenRowMask3 = 0;
                _seenRowCount = 0;
            }

            internal readonly bool AllRowsSeen => _rowCount > 0 && _seenRowCount == _rowCount;

            internal void MarkSeen(int rowOffset) {
                if ((uint)rowOffset >= (uint)_rowCount) {
                    return;
                }

                if (_seenRows == null
                    && _seenRowMask0 == 0
                    && _seenRowMask1 == 0
                    && _seenRowMask2 == 0
                    && _seenRowMask3 == 0) {
                    if (rowOffset < _seenRowCount) {
                        return;
                    }

                    if (rowOffset == _seenRowCount) {
                        _seenRowCount++;
                        return;
                    }

                    if (_rowCount > 256) {
                        _seenRows = CreateSeenRowsTracker(_seenRowCount, _rowCount);
                    } else {
                        MarkDensePrefixSeenInMasks(_seenRowCount, ref _seenRowMask0, ref _seenRowMask1, ref _seenRowMask2, ref _seenRowMask3);
                    }
                }

                if (_rowCount > 256) {
                    if (_seenRows == null) {
                        if (rowOffset < _seenRowCount) {
                            return;
                        }

                        if (rowOffset == _seenRowCount) {
                            _seenRowCount++;
                            return;
                        }

                        _seenRows = CreateSeenRowsTracker(_seenRowCount, _rowCount);
                    }

                    if (_seenRows[rowOffset]) {
                        return;
                    }

                    _seenRows[rowOffset] = true;
                    _seenRowCount++;
                    return;
                }

                if (_seenRows == null) {
                    int maskIndex = rowOffset >> 6;
                    ulong rowBit = 1UL << (rowOffset & 63);
                    switch (maskIndex) {
                        case 0:
                            if ((_seenRowMask0 & rowBit) != 0) {
                                return;
                            }

                            _seenRowMask0 |= rowBit;
                            break;
                        case 1:
                            if ((_seenRowMask1 & rowBit) != 0) {
                                return;
                            }

                            _seenRowMask1 |= rowBit;
                            break;
                        case 2:
                            if ((_seenRowMask2 & rowBit) != 0) {
                                return;
                            }

                            _seenRowMask2 |= rowBit;
                            break;
                        default:
                            if ((_seenRowMask3 & rowBit) != 0) {
                                return;
                            }

                            _seenRowMask3 |= rowBit;
                            break;
                    }
                } else {
                    if (_seenRows[rowOffset]) {
                        return;
                    }

                    _seenRows[rowOffset] = true;
                }

                _seenRowCount++;
            }
        }

        private static void MarkDensePrefixSeenInMasks(int seenDensePrefixLength, ref ulong mask0, ref ulong mask1, ref ulong mask2, ref ulong mask3) {
            if (seenDensePrefixLength <= 0) {
                return;
            }

            if (seenDensePrefixLength >= 64) {
                mask0 = ulong.MaxValue;
            } else {
                mask0 = (1UL << seenDensePrefixLength) - 1UL;
                return;
            }

            int remaining = seenDensePrefixLength - 64;
            if (remaining <= 0) {
                return;
            }

            if (remaining >= 64) {
                mask1 = ulong.MaxValue;
            } else {
                mask1 = (1UL << remaining) - 1UL;
                return;
            }

            remaining -= 64;
            if (remaining <= 0) {
                return;
            }

            if (remaining >= 64) {
                mask2 = ulong.MaxValue;
            } else {
                mask2 = (1UL << remaining) - 1UL;
                return;
            }

            remaining -= 64;
            if (remaining > 0) {
                mask3 = remaining >= 64 ? ulong.MaxValue : (1UL << remaining) - 1UL;
            }
        }

        private static bool[] CreateSeenRowsTracker(int seenDensePrefixLength, int rowCount) {
            var seenRows = new bool[rowCount];
            for (int i = 0; i < seenDensePrefixLength; i++) {
                seenRows[i] = true;
            }

            return seenRows;
        }

        private static ulong CreateAllColumnsSeenMask(int columnCount) {
            return columnCount == 64 ? ulong.MaxValue : (1UL << columnCount) - 1UL;
        }

        private static bool MarkRequestedColumnSeen(int columnOffset, ulong allColumnsSeen, ref ulong seenColumns) {
            seenColumns |= 1UL << columnOffset;
            return seenColumns == allColumnsSeen;
        }

        private void ReadXmlRowIntoRange(XmlReader rowReader, object?[,] result, int rowIndex, int r1, int c1, int c2, int width, object?[]? rowBuffer8, CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return;
            }

            int rr = rowIndex - r1;
            if ((uint)rr >= (uint)result.GetLength(0)) {
                SkipXmlElement(rowReader, "row");
                return;
            }

            if (width == 8 && rowBuffer8 != null) {
                ReadXmlRowIntoRange8(rowReader, result, rr, c1, c2, rowBuffer8, ct);
                return;
            }

            if (width == 3) {
                ReadXmlRowIntoRange3(rowReader, result, rr, c1, c2, ct);
                return;
            }

            if (width == 10) {
                ReadXmlRowIntoRangeKnownWidth(rowReader, result, rr, c1, c2, width, 0x3FFUL, ct);
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            bool canTrackColumns = width <= 64;
            ulong allColumnsSeen = canTrackColumns ? CreateAllColumnsSeenMask(width) : 0UL;
            ulong seenColumns = 0;
            bool canUseOrderedFullWidthExit = canTrackColumns;
            int nextExpectedColumn = c1;
            int visitedNodes = 0;
            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    if (canUseOrderedFullWidthExit) {
                        canUseOrderedFullWidthExit = false;
                        int orderedSeen = nextExpectedColumn - c1;
                        seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    }

                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex < c1 || columnIndex > c2) {
                    if (canUseOrderedFullWidthExit && columnIndex > c2 && nextExpectedColumn <= c2) {
                        canUseOrderedFullWidthExit = false;
                        int orderedSeen = nextExpectedColumn - c1;
                        seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    }

                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int cc = columnIndex - c1;
                if ((uint)cc >= (uint)width) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    canUseOrderedFullWidthExit = false;
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                }

                result[rr, cc] = ReadXmlCellValue(rowReader, rowReader.GetAttribute("t"));
                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                }

                if (canUseOrderedFullWidthExit && columnIndex >= c2) {
                    SkipXmlElementContent(rowReader, depth);
                    return;
                }

                if (canTrackColumns && !canUseOrderedFullWidthExit && MarkRequestedColumnSeen(cc, allColumnsSeen, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth);
                    return;
                }
            }
        }

        private void ReadXmlRowIntoRange8(XmlReader rowReader, object?[,] result, int rowOffset, int c1, int c2, object?[] rowBuffer, CancellationToken ct) {
            byte seenColumns = ReadXmlRowIntoChunk8(rowReader, rowBuffer, c1, c2, ct);
            if (seenColumns == 0xFF) {
                result[rowOffset, 0] = rowBuffer[0];
                result[rowOffset, 1] = rowBuffer[1];
                result[rowOffset, 2] = rowBuffer[2];
                result[rowOffset, 3] = rowBuffer[3];
                result[rowOffset, 4] = rowBuffer[4];
                result[rowOffset, 5] = rowBuffer[5];
                result[rowOffset, 6] = rowBuffer[6];
                result[rowOffset, 7] = rowBuffer[7];
                return;
            }

            if ((seenColumns & 0x01) != 0) result[rowOffset, 0] = rowBuffer[0];
            if ((seenColumns & 0x02) != 0) result[rowOffset, 1] = rowBuffer[1];
            if ((seenColumns & 0x04) != 0) result[rowOffset, 2] = rowBuffer[2];
            if ((seenColumns & 0x08) != 0) result[rowOffset, 3] = rowBuffer[3];
            if ((seenColumns & 0x10) != 0) result[rowOffset, 4] = rowBuffer[4];
            if ((seenColumns & 0x20) != 0) result[rowOffset, 5] = rowBuffer[5];
            if ((seenColumns & 0x40) != 0) result[rowOffset, 6] = rowBuffer[6];
            if ((seenColumns & 0x80) != 0) result[rowOffset, 7] = rowBuffer[7];
        }

        private void ReadXmlRowIntoRange3(XmlReader rowReader, object?[,] result, int rowOffset, int c1, int c2, CancellationToken ct) {
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int nextExpectedColumn = c1;
            bool canUseOrderedFullWidthExit = true;
            ulong seenColumns = 0;
            int visitedNodes = 0;

            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    if (canUseOrderedFullWidthExit) {
                        int orderedSeen = nextExpectedColumn - c1;
                        seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    }

                    canUseOrderedFullWidthExit = false;
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex < c1 || columnIndex > c2) {
                    if (canUseOrderedFullWidthExit && columnIndex > c2 && nextExpectedColumn <= c2) {
                        int orderedSeen = nextExpectedColumn - c1;
                        seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                        canUseOrderedFullWidthExit = false;
                    }

                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int columnOffset = columnIndex - c1;
                if ((uint)columnOffset >= 3U) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    canUseOrderedFullWidthExit = false;
                }

                result[rowOffset, columnOffset] = ReadXmlCellValue(rowReader, rowReader.GetAttribute("t"));

                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                    if (columnIndex >= c2) {
                        SkipXmlElementContent(rowReader, depth);
                        return;
                    }
                } else {
                    seenColumns |= 1UL << columnOffset;
                    if (seenColumns == 0x7UL) {
                        SkipXmlElementContent(rowReader, depth);
                        return;
                    }
                }
            }
        }

        private void ReadXmlRowIntoRangeKnownWidth(XmlReader rowReader, object?[,] result, int rowOffset, int c1, int c2, int width, ulong allColumnsSeen, CancellationToken ct) {
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int nextExpectedColumn = c1;
            bool canUseOrderedFullWidthExit = true;
            ulong seenColumns = 0;
            int visitedNodes = 0;

            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    if (canUseOrderedFullWidthExit) {
                        int orderedSeen = nextExpectedColumn - c1;
                        seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    }

                    canUseOrderedFullWidthExit = false;
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex < c1 || columnIndex > c2) {
                    if (canUseOrderedFullWidthExit && columnIndex > c2 && nextExpectedColumn <= c2) {
                        int orderedSeen = nextExpectedColumn - c1;
                        seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                        canUseOrderedFullWidthExit = false;
                    }

                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int columnOffset = columnIndex - c1;
                if ((uint)columnOffset >= (uint)width) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    canUseOrderedFullWidthExit = false;
                }

                result[rowOffset, columnOffset] = ReadXmlCellValue(rowReader, rowReader.GetAttribute("t"));

                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                    if (columnIndex >= c2) {
                        SkipXmlElementContent(rowReader, depth);
                        return;
                    }
                } else {
                    seenColumns |= 1UL << columnOffset;
                    if (seenColumns == allColumnsSeen) {
                        SkipXmlElementContent(rowReader, depth);
                        return;
                    }
                }
            }
        }

    }
}
