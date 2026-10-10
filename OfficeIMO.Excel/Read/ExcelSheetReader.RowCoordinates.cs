#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private Dictionary<long, int>? _implicitXmlRowIndexes;

        // A completed worksheet scan may populate this index before a streaming
        // projection needs it. Direct range readers retain the dedicated scan.
        private int ResolveImplicitXmlRowIndex(XmlReader reader, int fallback, CancellationToken ct) {
            ct.ThrowIfCancellationRequested();
            long position = GetXmlRowPosition(reader);
            var indexes = Volatile.Read(ref _implicitXmlRowIndexes);
            if (indexes == null) {
                var discovered = ReadImplicitXmlRowIndexes(ct);
                indexes = Interlocked.CompareExchange(ref _implicitXmlRowIndexes, discovered, null) ?? discovered;
            }
            return indexes.TryGetValue(position, out int rowIndex) ? rowIndex : fallback;
        }

        private Dictionary<long, int> ReadImplicitXmlRowIndexes(CancellationToken ct) {
            using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
            RewindWorksheetStream(stream);
            using var reader = OpenWorksheetXmlReader(stream);
            var coordinates = new ImplicitXmlRowIndexBuilder();
            var worksheetRows = new WorksheetXmlRowSelector();
            int rowDepth = -1;
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                if (worksheetRows.IsRowElement(reader)) {
                    coordinates.BeginRow(reader, ParsePositiveIntAttribute(ReadXmlReferenceAttribute(reader).Text));
                    if (reader.IsEmptyElement) coordinates.EndRow();
                    else rowDepth = reader.Depth;
                } else if (rowDepth >= 0 && SpreadsheetXmlContent.IsDirectChildElement(reader, rowDepth, "c")) {
                    if (coordinates.NeedsCellReference) coordinates.AddCell(ReadXmlReferenceAttribute(reader).Text);
                } else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row" && reader.Depth == rowDepth) {
                    coordinates.EndRow();
                    rowDepth = -1;
                }
            }
            ct.ThrowIfCancellationRequested();
            return coordinates.Indexes;
        }

        // Shared by full validation and the fallback scan. It retains only rows
        // whose first cell reference moves an omitted row index away from its
        // sequential position, never the cells or their values.
        private sealed class ImplicitXmlRowIndexBuilder {
            internal Dictionary<long, int> Indexes { get; } = new Dictionary<long, int>();
            private int _nextRowIndex = 1;
            private int _rowIndex;
            private int _fallback;
            private long _position;
            private bool _inferred;
            private bool _hasPreviousRow;
            private int _previousRowIndex;
            internal bool NeedsCellReference { get; private set; }
            internal bool RowsStrictlyIncreasing { get; private set; } = true;

            internal void BeginRow(XmlReader reader, int declaredRowIndex) {
                _rowIndex = declaredRowIndex;
                _inferred = _rowIndex <= 0;
                _fallback = _nextRowIndex;
                if (_inferred) _rowIndex = _fallback;
                NeedsCellReference = _inferred;
                _position = _inferred ? GetXmlRowPosition(reader) : 0;
            }

            internal void AddCell(ReadOnlySpan<char> reference) {
                if (NeedsCellReference && A1.TryParseCellReferenceFast(reference, out int rowIndex, out _)) {
                    _rowIndex = rowIndex;
                    NeedsCellReference = false;
                }
            }

            internal void EndRow() {
                if (_hasPreviousRow && _rowIndex <= _previousRowIndex) RowsStrictlyIncreasing = false;
                _previousRowIndex = _rowIndex;
                _hasPreviousRow = true;
                if (_inferred && _rowIndex != _fallback) {
                    if (Indexes.Count >= A1.MaxRows) {
                        throw new InvalidDataException("Worksheet implicit row coordinates exceed the XLSX row limit.");
                    }
                    Indexes.Add(_position, _rowIndex);
                }
                _nextRowIndex = _rowIndex + 1;
                NeedsCellReference = false;
            }
        }

        private static long GetXmlRowPosition(XmlReader reader) {
            if (reader is not IXmlLineInfo info || !info.HasLineInfo()) {
                throw new InvalidDataException("Worksheet XML reader does not expose row positions.");
            }
            return ((long)(uint)info.LineNumber << 32) | (uint)info.LinePosition;
        }
    }
}
