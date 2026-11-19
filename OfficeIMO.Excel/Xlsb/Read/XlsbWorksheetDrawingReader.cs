using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Decodes the worksheet BrtDrawing record (MS-XLSB 2.4.354), the link to the sheet's DrawingML drawing part.</summary>
    internal sealed class XlsbWorksheetDrawingReader {
        private readonly XlsbWorksheetFeatureReader _owner;

        internal XlsbWorksheetDrawingReader(XlsbWorksheetFeatureReader owner) => _owner = owner;

        internal bool TryRead(XlsbRecord record) {
            if (record.Type != XlsbWorksheetDrawingWriter.BrtDrawing) return false;
            if (_owner.Sheet.DrawingRelationshipId != null) {
                throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains more than one BrtDrawing record.");
            }
            var cursor = new XlsbBinaryCursor(record.Data);
            uint length = cursor.ReadUInt32();
            // RelID is an XLNullableWideString of at most 255 characters; NULL means no drawing is linked.
            if (length == 0xFFFFFFFFU) {
                if (cursor.Remaining != 0) throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an invalid BrtDrawing record.");
                return true;
            }
            if (length == 0 || length > 255 || cursor.Remaining != length * 2L) {
                throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an invalid BrtDrawing relationship id.");
            }
            var characters = new char[length];
            for (int index = 0; index < characters.Length; index++) {
                characters[index] = (char)cursor.ReadUInt16();
                if (characters[index] == '\0') throw new InvalidDataException($"The XLSB worksheet part '{_owner.PartName}' contains an invalid BrtDrawing relationship id.");
            }
            _owner.Sheet.DrawingRelationshipId = new string(characters);
            return true;
        }
    }
}
