using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Streaming APIs for large ranges.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private void ReadXmlRowIntoChunk(XmlReader rowReader, object?[][] rows, int rowIndex, int startRow, int c1, int c2, CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return;
            }

            int rowOffset = rowIndex - startRow;
            if ((uint)rowOffset >= (uint)rows.Length) {
                return;
            }

            object?[] rowValues = rows[rowOffset];
            if (rowValues.Length == 8) {
                ReadXmlRowIntoChunk8(rowReader, rowValues, c1, c2, ct);
                return;
            }

            if (rowValues.Length == 3) {
                ReadXmlRowIntoChunkKnownWidth(rowReader, rowValues, c1, c2, 3, 0x7UL, ct);
                return;
            }

            if (rowValues.Length == 10) {
                ReadXmlRowIntoChunkKnownWidth(rowReader, rowValues, c1, c2, 10, 0x3FFUL, ct);
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            bool canTrackColumns = rowValues.Length <= 64;
            ulong allColumnsSeen = canTrackColumns ? CreateAllColumnsSeenMask(rowValues.Length) : 0UL;
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

                int columnOffset = columnIndex - c1;
                if ((uint)columnOffset >= (uint)rowValues.Length) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    canUseOrderedFullWidthExit = false;
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                }

                rowValues[columnOffset] = ReadXmlCellValue(rowReader, rowReader.GetAttribute("t"));
                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                }

                if (canUseOrderedFullWidthExit && columnIndex >= c2) {
                    SkipXmlElementContent(rowReader, depth);
                    return;
                }

                if (canTrackColumns && !canUseOrderedFullWidthExit && MarkRequestedColumnSeen(columnOffset, allColumnsSeen, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth);
                    return;
                }
            }
        }

        private byte ReadXmlRowIntoChunk8(XmlReader rowReader, object?[] rowValues, int c1, int c2, CancellationToken ct) {
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int nextExpectedColumn = c1;
            bool canUseOrderedFullWidthExit = true;
            ulong seenColumns = 0;
            byte seenColumnMask = 0;
            int visitedNodes = 0;

            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return seenColumnMask;
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
                if ((uint)columnOffset >= 8U) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    canUseOrderedFullWidthExit = false;
                }

                rowValues[columnOffset] = ReadXmlCellValue(rowReader, rowReader.GetAttribute("t"));
                seenColumnMask |= (byte)(1 << columnOffset);

                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                    if (columnIndex >= c2) {
                        SkipXmlElementContent(rowReader, depth);
                        return seenColumnMask;
                    }
                } else {
                    seenColumns |= 1UL << columnOffset;
                    if (seenColumns == 0xFFUL) {
                        SkipXmlElementContent(rowReader, depth);
                        return seenColumnMask;
                    }
                }
            }

            return seenColumnMask;
        }

        private void ReadXmlRowIntoChunkKnownWidth(XmlReader rowReader, object?[] rowValues, int c1, int c2, int width, ulong allColumnsSeen, CancellationToken ct) {
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

                rowValues[columnOffset] = ReadXmlCellValue(rowReader, rowReader.GetAttribute("t"));

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