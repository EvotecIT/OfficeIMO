using System.IO;
using System.Threading;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelSheetReader {
        internal uint?[,]? ReadCellStyleIndexes(string a1Range, CancellationToken ct = default) {
            var (firstRow, firstColumn, lastRow, lastColumn) = A1.ParseRange(a1Range);
            if (!CanAttemptXmlFastReader()) {
                return null;
            }

            int rowCount = lastRow - firstRow + 1;
            int columnCount = lastColumn - firstColumn + 1;
            var styles = new uint?[rowCount, columnCount];
            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                int nextRowIndex = 1;
                while (reader.Read()) {
                    ct.ThrowIfCancellationRequested();
                    if (reader.NodeType != System.Xml.XmlNodeType.Element || reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                    if (rowIndex <= 0) {
                        rowIndex = nextRowIndex;
                    }
                    nextRowIndex = rowIndex + 1;

                    if (rowIndex > lastRow) {
                        break;
                    }
                    if (rowIndex < firstRow || reader.IsEmptyElement) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    int rowDepth = reader.Depth;
                    int nextColumnIndex = 1;
                    bool advanceReader = true;
                    while (advanceReader ? reader.Read() : !reader.EOF) {
                        advanceReader = true;
                        if (reader.NodeType == System.Xml.XmlNodeType.EndElement &&
                            reader.Depth == rowDepth &&
                            reader.LocalName == "row") {
                            break;
                        }
                        if (reader.NodeType != System.Xml.XmlNodeType.Element || reader.LocalName != "c") {
                            continue;
                        }

                        int columnIndex = GetXmlCellColumnIndex(reader, ref nextColumnIndex);
                        if (columnIndex >= firstColumn && columnIndex <= lastColumn) {
                            uint styleIndex = TryParseUInt(reader.GetAttribute("s"), out uint parsedStyle)
                                ? parsedStyle
                                : 0U;
                            styles[rowIndex - firstRow, columnIndex - firstColumn] = styleIndex;
                        }

                        if (!reader.IsEmptyElement) {
                            reader.Skip();
                            advanceReader = false;
                        }
                    }
                }

                return styles;
            } catch (System.Xml.XmlException) {
                return null;
            } catch (IOException) {
                return null;
            } catch (UnauthorizedAccessException) {
                return null;
            } catch (ObjectDisposedException) {
                return null;
            }
        }
    }
}
