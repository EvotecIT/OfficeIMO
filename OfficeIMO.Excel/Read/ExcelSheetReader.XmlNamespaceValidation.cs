using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private enum IndexedXmlNamespaceStatus { Accepted, RequiresSdk, UnexpectedNamespace }

        /// <summary>Tracks the namespace and ancestry invariants required by the byte index.</summary>
        private sealed class IndexedWorksheetNamespaceValidator {
            private readonly bool _strictOuterShape;
            private int _sheetDataDepth = -1;
            private int _rowDepth = -1;
            private int _cellDepth = -1;
            private int _inlineStringDepth = -1;
            private bool _indexedSheetDataCompleted;

            internal IndexedWorksheetNamespaceValidator(bool strictOuterShape) => _strictOuterShape = strictOuterShape;

            internal IndexedXmlNamespaceStatus Observe(XmlReader reader) {
                if (reader.NodeType == XmlNodeType.Element) {
                    bool spreadsheetElement = string.Equals(reader.NamespaceURI, SpreadsheetNamespace, StringComparison.Ordinal)
                        || string.Equals(reader.NamespaceURI, StrictSpreadsheetNamespace, StringComparison.Ordinal);
                    if (_strictOuterShape) {
                        if (reader.Depth == 0 && (reader.LocalName != "worksheet" || !spreadsheetElement))
                            return IndexedXmlNamespaceStatus.RequiresSdk;
                        if ((reader.LocalName == "dimension" || reader.LocalName == "sheetData")
                            && (reader.Depth != 1 || !spreadsheetElement))
                            return IndexedXmlNamespaceStatus.RequiresSdk;
                    }
                    if (reader.LocalName == "sheetData") {
                        if (_indexedSheetDataCompleted) return IndexedXmlNamespaceStatus.Accepted;
                        if (!spreadsheetElement) return IndexedXmlNamespaceStatus.UnexpectedNamespace;
                        if (reader.IsEmptyElement) _indexedSheetDataCompleted = true;
                        else _sheetDataDepth = reader.Depth;
                        return IndexedXmlNamespaceStatus.Accepted;
                    }
                    if (_sheetDataDepth >= 0 && reader.Depth == _sheetDataDepth + 1 && reader.LocalName == "row") {
                        if (!spreadsheetElement) return IndexedXmlNamespaceStatus.UnexpectedNamespace;
                        _rowDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                        return IndexedXmlNamespaceStatus.Accepted;
                    }
                    if (_rowDepth >= 0 && reader.Depth == _rowDepth + 1 && reader.LocalName == "c") {
                        if (!spreadsheetElement) return IndexedXmlNamespaceStatus.UnexpectedNamespace;
                        _cellDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                        _inlineStringDepth = -1;
                        return IndexedXmlNamespaceStatus.Accepted;
                    }
                    if (_cellDepth >= 0 && reader.Depth == _cellDepth + 1) {
                        if (reader.LocalName == "v" || reader.LocalName == "f" || reader.LocalName == "is") {
                            if (!spreadsheetElement) return IndexedXmlNamespaceStatus.UnexpectedNamespace;
                            if (reader.LocalName == "is" && !reader.IsEmptyElement) _inlineStringDepth = reader.Depth;
                        }
                        return IndexedXmlNamespaceStatus.Accepted;
                    }
                    if (_inlineStringDepth >= 0 && reader.Depth == _inlineStringDepth + 1
                        && reader.LocalName == "t" && !spreadsheetElement)
                        return IndexedXmlNamespaceStatus.UnexpectedNamespace;
                } else if (reader.NodeType == XmlNodeType.EndElement) {
                    if (reader.Depth == _inlineStringDepth && reader.LocalName == "is") _inlineStringDepth = -1;
                    if (reader.Depth == _cellDepth && reader.LocalName == "c") {
                        _cellDepth = -1;
                        _inlineStringDepth = -1;
                    }
                    if (reader.Depth == _rowDepth && reader.LocalName == "row") _rowDepth = -1;
                    if (reader.Depth == _sheetDataDepth && reader.LocalName == "sheetData") {
                        _sheetDataDepth = -1;
                        _indexedSheetDataCompleted = true;
                    }
                }
                return IndexedXmlNamespaceStatus.Accepted;
            }
        }
    }
}
