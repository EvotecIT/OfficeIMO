#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private Dictionary<long, int>? _implicitXmlRowIndexes;

        // Explicit row indices stay on the ordinary path. Missing indices need a
        // lookahead because their first referenced cell may follow a sparse gap.
        // A separate scan preserves the caller's XmlReader position and caches
        // only deviations from sequential numbering, never cell values or rows.
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
            var indexes = new Dictionary<long, int>();
            int nextRowIndex = 1;
            int rowDepth = -1;
            int rowIndex = 0;
            int fallback = 0;
            long position = 0;
            bool inferred = false;
            bool hasCellReference = false;
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row" && rowDepth < 0) {
                    rowIndex = ParsePositiveIntAttribute(ReadXmlReferenceAttribute(reader).Text);
                    inferred = rowIndex <= 0;
                    fallback = nextRowIndex;
                    if (inferred) rowIndex = fallback;
                    hasCellReference = false;
                    position = inferred ? GetXmlRowPosition(reader) : 0;
                    if (reader.IsEmptyElement) {
                        nextRowIndex = rowIndex + 1;
                    } else {
                        rowDepth = reader.Depth;
                    }
                } else if (inferred && !hasCellReference && rowDepth >= 0
                    && reader.NodeType == XmlNodeType.Element && reader.LocalName == "c" && reader.Depth == rowDepth + 1) {
                    if (A1.TryParseCellReferenceFast(ReadXmlReferenceAttribute(reader).Text, out int referencedRow, out _)) {
                        rowIndex = referencedRow;
                        hasCellReference = true;
                    }
                } else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row" && reader.Depth == rowDepth) {
                    if (inferred && rowIndex != fallback) {
                        if (indexes.Count >= A1.MaxRows) {
                            throw new InvalidDataException("Worksheet implicit row coordinates exceed the XLSX row limit.");
                        }
                        indexes.Add(position, rowIndex);
                    }
                    nextRowIndex = rowIndex + 1;
                    rowDepth = -1;
                }
            }
            // Do not publish a partially scanned or cancelled index.
            ct.ThrowIfCancellationRequested();
            return indexes;
        }

        private static long GetXmlRowPosition(XmlReader reader) {
            if (reader is not IXmlLineInfo info || !info.HasLineInfo()) {
                throw new InvalidDataException("Worksheet XML reader does not expose row positions.");
            }
            return ((long)(uint)info.LineNumber << 32) | (uint)info.LinePosition;
        }
    }
}
