using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Decodes the worksheet LISTPARTS collection (MS-XLSB 2.4.103, 2.4.707) into tableParts.</summary>
    internal sealed class XlsbWorksheetTableReader {
        private readonly XlsbWorksheetFeatureReader _owner;
        private TableParts? _current;
        private uint _declared;
        private bool _seen;
        private bool _unresolved;

        internal XlsbWorksheetTableReader(XlsbWorksheetFeatureReader owner) => _owner = owner;

        internal bool TryRead(XlsbRecord record) {
            switch (record.Type) {
                case XlsbWorksheetTableWriter.BrtBeginListParts:
                    if (_current != null || _seen || record.Data.Length != 4) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an invalid or repeated table-part collection.");
                    }
                    _seen = true;
                    _declared = BitConverter.ToUInt32(record.Data, 0);
                    _current = new TableParts();
                    return true;
                case XlsbWorksheetTableWriter.BrtListPart:
                    if (_current == null) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains a BrtListPart record outside its collection.");
                    }
                    var cursor = new XlsbBinaryCursor(record.Data);
                    uint count = cursor.ReadUInt32();
                    if (count == 0xFFFFFFFFU || count == 0U || count > 255U) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains a BrtListPart record without a relationship id.");
                    }
                    string id = Encoding.Unicode.GetString(cursor.ReadBytes(checked((int)count * 2)));
                    if (cursor.Remaining != 0 || _current.Elements<TablePart>().Any(part => part.Id?.Value == id)) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an invalid or repeated table relationship '{id}'.");
                    }
                    // A part that is not a decodable table keeps the whole collection native and unprojected.
                    if (!_owner.Sheet.Tables.ContainsKey(id)) _unresolved = true;
                    _current.Append(new TablePart { Id = id });
                    return true;
                case XlsbWorksheetTableWriter.BrtEndListParts:
                    if (_current == null) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an unmatched table-part end record.");
                    }
                    uint actual = (uint)_current.ChildElements.Count;
                    if (actual != _declared) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' declares {_declared} table parts but lists {actual}.");
                    }
                    // OfficeIMO states the derived count when it adds tables; project it the same way.
                    _current.Count = actual;
                    if (_unresolved) {
                        _owner.Workbook.AddDiagnostic(new XlsbImportDiagnostic("xlsb-table-partial", XlsbImportDiagnosticSeverity.Warning,
                            $"Kept the tables of worksheet '{_owner.Sheet.Name}' native because a listed part is not a readable table.", _owner.PartName));
                    } else if (actual != 0) {
                        _owner.Sheet.AddFeatureElement(_current);
                    }
                    _current = null;
                    return true;
                default:
                    return false;
            }
        }

        internal void Complete() {
            if (_current != null) {
                throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an unterminated table-part collection.");
            }
        }
    }
}
