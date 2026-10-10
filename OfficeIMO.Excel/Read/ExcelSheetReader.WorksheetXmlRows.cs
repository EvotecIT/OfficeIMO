using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        /// <summary>
        /// Tracks the actual worksheet, sheetData and row owners while scanning XML.
        /// Row helpers may consume a subtree; visiting its next peer or ancestor closes
        /// stale state without retaining XML nodes or allocating per scanned element.
        /// </summary>
        private struct WorksheetXmlRowSelector {
            private bool _hasWorksheetRoot;
            private bool _inSheetData;
            private bool _inRow;
            private int _worksheetDepth;
            private int _sheetDataDepth;
            private int _rowDepth;

            internal bool IsRowElement(XmlReader reader) {
                Observe(reader);
                return _inSheetData && SpreadsheetXmlContent.IsDirectChildElement(reader, _sheetDataDepth, "row");
            }

            internal bool IsCellElement(XmlReader reader) {
                Observe(reader);
                return _inRow && SpreadsheetXmlContent.IsDirectChildElement(reader, _rowDepth, "c");
            }

            internal bool IsWorksheetChildElement(XmlReader reader, string localName) {
                Observe(reader);
                return _hasWorksheetRoot && SpreadsheetXmlContent.IsDirectChildElement(reader, _worksheetDepth, localName);
            }

            private void Observe(XmlReader reader) {
                if (_inRow && reader.Depth <= _rowDepth) _inRow = false;
                if (_inSheetData && reader.Depth <= _sheetDataDepth) {
                    _inSheetData = false;
                    _inRow = false;
                }

                if (reader.NodeType != XmlNodeType.Element) {
                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth <= _worksheetDepth) {
                        _hasWorksheetRoot = false;
                    }
                    return;
                }

                if (reader.Depth == 0) {
                    _hasWorksheetRoot = SpreadsheetXmlContent.IsSpreadsheetElement(reader, "worksheet");
                    _worksheetDepth = reader.Depth;
                    _inSheetData = false;
                    _inRow = false;
                    return;
                }

                if (!_hasWorksheetRoot) return;
                if (SpreadsheetXmlContent.IsDirectChildElement(reader, _worksheetDepth, "sheetData")) {
                    _sheetDataDepth = reader.Depth;
                    _inSheetData = !reader.IsEmptyElement;
                    return;
                }

                if (_inSheetData && SpreadsheetXmlContent.IsDirectChildElement(reader, _sheetDataDepth, "row")) {
                    _rowDepth = reader.Depth;
                    _inRow = !reader.IsEmptyElement;
                }
            }
        }
    }
}
