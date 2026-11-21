using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Decodes RWBRK and COLBRK collections (MS-XLSB 2.4.197, 2.4.25, 2.4.315) into rowBreaks and colBreaks.</summary>
    internal sealed class XlsbWorksheetPageBreakReader {
        private const int BrtBeginRwBrk = 392;
        private const int BrtEndRwBrk = 393;
        private const int BrtBeginColBrk = 394;
        private const int BrtEndColBrk = 395;
        private const int BrtBrk = 396;

        private readonly XlsbWorksheetFeatureReader _owner;
        private OpenXmlCompositeElement? _current;
        private bool _sawRows;
        private bool _sawColumns;

        internal XlsbWorksheetPageBreakReader(XlsbWorksheetFeatureReader owner) => _owner = owner;

        internal bool TryRead(XlsbRecord record) {
            switch (record.Type) {
                case BrtBeginRwBrk:
                case BrtBeginColBrk:
                    bool rows = record.Type == BrtBeginRwBrk;
                    if (_current != null || (rows ? _sawRows : _sawColumns) || record.Data.Length != 8) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an invalid or repeated page-break collection.");
                    }
                    if (rows) _sawRows = true; else _sawColumns = true;
                    _current = rows ? new RowBreaks() : new ColumnBreaks();
                    return true;
                case BrtBrk:
                    if (_current == null || record.Data.Length != 20) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains a BrtBrk record outside a page-break collection.");
                    }
                    var item = new Break {
                        Id = BitConverter.ToUInt32(record.Data, 0),
                        Min = BitConverter.ToUInt32(record.Data, 4),
                        Max = BitConverter.ToUInt32(record.Data, 8),
                    };
                    if (BitConverter.ToUInt32(record.Data, 12) != 0) item.ManualPageBreak = true;
                    if (BitConverter.ToUInt32(record.Data, 16) != 0) item.PivotTablePageBreak = true;
                    _current.Append(item);
                    return true;
                case BrtEndRwBrk:
                case BrtEndColBrk:
                    if (_current == null || (record.Type == BrtEndRwBrk) != (_current is RowBreaks)) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an unmatched page-break end record.");
                    }
                    Break[] breaks = _current.Elements<Break>().ToArray();
                    uint manual = (uint)breaks.Count(item => item.ManualPageBreak?.Value == true);
                    if (_current is RowBreaks rowBreaks) {
                        rowBreaks.Count = (uint)breaks.Length;
                        rowBreaks.ManualBreakCount = manual;
                    } else {
                        var columnBreaks = (ColumnBreaks)_current;
                        columnBreaks.Count = (uint)breaks.Length;
                        columnBreaks.ManualBreakCount = manual;
                    }
                    if (breaks.Length != 0) _owner.Sheet.AddFeatureElement(_current);
                    _current = null;
                    return true;
                default:
                    return false;
            }
        }

        internal void Complete() {
            if (_current != null) {
                throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an unterminated page-break collection.");
            }
        }
    }
}
