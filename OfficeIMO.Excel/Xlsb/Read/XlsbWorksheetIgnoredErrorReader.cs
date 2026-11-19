using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Decodes an IGNOREECS collection (MS-XLSB 2.4.16, 2.4.322) into ignoredErrors.</summary>
    internal sealed class XlsbWorksheetIgnoredErrorReader {
        private readonly XlsbWorksheetFeatureReader _owner;
        private IgnoredErrors? _current;
        private bool _seen;

        internal XlsbWorksheetIgnoredErrorReader(XlsbWorksheetFeatureReader owner) => _owner = owner;

        internal bool TryRead(XlsbRecord record) {
            switch (record.Type) {
                case XlsbWorksheetIgnoredErrorWriter.BrtBeginCellIgnoreECs:
                    if (_current != null || _seen) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains a repeated ignored-error collection.");
                    }
                    _seen = true;
                    _current = new IgnoredErrors();
                    return true;
                case XlsbWorksheetIgnoredErrorWriter.BrtCellIgnoreEC:
                    if (_current == null || record.Data.Length < 8) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains a BrtCellIgnoreEC record outside its collection.");
                    }
                    uint flags = BitConverter.ToUInt32(record.Data, 0);
                    int offset = 4;
                    var item = new IgnoredError {
                        SequenceOfReferences = new ListValue<StringValue> {
                            InnerText = XlsbWorksheetFeatureReader.ReadSqRfX(record.Data, ref offset, _owner.PartName)
                        }
                    };
                    foreach (var entry in XlsbWorksheetIgnoredErrorWriter.Flags) {
                        if ((flags & entry.Flag) != 0) entry.Set(item);
                    }
                    _current.Append(item);
                    return true;
                case XlsbWorksheetIgnoredErrorWriter.BrtEndCellIgnoreECs:
                    if (_current == null) {
                        throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an unmatched ignored-error end record.");
                    }
                    // OfficeIMO states the derived count on save; project it the same way so unchanged content compares equal.
                    int count = _current.Elements<IgnoredError>().Count();
                    _current.SetAttribute(new OpenXmlAttribute("count", string.Empty, count.ToString(System.Globalization.CultureInfo.InvariantCulture)));
                    if (count != 0) _owner.Sheet.AddFeatureElement(_current);
                    _current = null;
                    return true;
                default:
                    return false;
            }
        }

        internal void Complete() {
            if (_current != null) {
                throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an unterminated ignored-error collection.");
            }
        }
    }
}
