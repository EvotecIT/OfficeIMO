using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Column-oriented readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Reads a single-column A1 range (e.g., "B2:B1000") as a typed sequence.
        /// </summary>
        public IEnumerable<object?> ReadColumn(string a1Range, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (c1 != c2) throw new ArgumentException("ReadColumn expects a single-column A1 range (e.g., 'B2:B100').", nameof(a1Range));

            bool canCancel = ct.CanBeCanceled;
            int height = r2 - r1 + 1;
            if (CanUseColumnXmlReader()) {
                if (TryReadColumnXmlFast(r1, c1, r2, height, ct, out var xmlValues)) {
                    for (int i = 0; i < height; i++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        yield return xmlValues[i];
                    }

                    yield break;
                }
            }

            if (height > DenseSnapshotCapacityLimit && RowsAreSortedWithinRange(r1, r2, ct)) {
                int nextRow = r1;
                foreach (var row in EnumerateWorksheetRows(ct)) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int ri = checked((int)row.RowIndex!.Value);
                    if (ri < r1) continue;
                    if (ri > r2) break;

                    while (nextRow < ri) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        yield return null;
                        nextRow++;
                    }

                    yield return ReadColumnValue(row, c1, ct);
                    nextRow = ri + 1;
                }

                while (nextRow <= r2) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    yield return null;
                    nextRow++;
                }

                yield break;
            }

            var rowMap = new Dictionary<int, Row>(GetSnapshotCapacity(height));
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int ri = checked((int)row.RowIndex!.Value);
                if (ri < r1) continue;
                if (ri > r2) continue;
                rowMap[ri] = row;
            }

            for (int r = r1; r <= r2; r++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (!rowMap.TryGetValue(r, out var row)) { yield return null; continue; }

                yield return ReadColumnValue(row, c1, ct);
            }

            object? ReadColumnValue(Row row, int columnIndex, CancellationToken token) {
                bool canCancelCell = token.CanBeCanceled;
                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancelCell) {
                        token.ThrowIfCancellationRequested();
                    }

                    int cc = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cc != columnIndex) continue;
                    return TryConvertCell(cell, out object? value) ? value : null;
                }

                return null;
            }
        }

        private bool TryReadColumnXmlFast(int r1, int columnIndex, int r2, int height, CancellationToken ct, out object?[] values) {
            values = new object?[height];

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                var seenRows = CreateCompletedRowTracker(height);
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

                    int rowOffset = rowIndex - r1;
                    if ((uint)rowOffset >= (uint)values.Length) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    ReadXmlColumnValue(reader, values, rowOffset, columnIndex, ct);
                    seenRows.MarkSeen(rowOffset);
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

        private void ReadXmlColumnValue(XmlReader rowReader, object?[] values, int rowOffset, int targetColumnIndex, CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            while (rowReader.Read()) {
                if (canCancel) {
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
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex != targetColumnIndex) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                values[rowOffset] = ReadXmlCellValue(rowReader);
                SkipXmlElementContent(rowReader, depth, "row");
                return;
            }
        }

        private bool CanUseColumnXmlReader() {
            return (_opt.CellValueConverter != null || _opt.Culture == System.Globalization.CultureInfo.InvariantCulture)
                && CanStreamWorksheetPart();
        }
    }
}

