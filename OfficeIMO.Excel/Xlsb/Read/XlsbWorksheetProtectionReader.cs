using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Decodes classic BIFF12 worksheet-protection settings.</summary>
    internal static class XlsbWorksheetProtectionReader {
        internal static XlsbWorksheetProtection Read(XlsbRecord record) {
            if (record.Data.Length != 66) {
                throw new InvalidDataException($"The BrtSheetProtection record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }
            var cursor = new XlsbBinaryCursor(record.Data);
            return new XlsbWorksheetProtection(
                cursor.ReadUInt16(),
                ReadBoolean(cursor, record, "fLocked"),
                ReadBoolean(cursor, record, "fObjects"),
                ReadBoolean(cursor, record, "fScenarios"),
                ReadBoolean(cursor, record, "fFormatCells"),
                ReadBoolean(cursor, record, "fFormatColumns"),
                ReadBoolean(cursor, record, "fFormatRows"),
                ReadBoolean(cursor, record, "fInsertColumns"),
                ReadBoolean(cursor, record, "fInsertRows"),
                ReadBoolean(cursor, record, "fInsertHyperlinks"),
                ReadBoolean(cursor, record, "fDeleteColumns"),
                ReadBoolean(cursor, record, "fDeleteRows"),
                ReadBoolean(cursor, record, "fSelLockedCells"),
                ReadBoolean(cursor, record, "fSort"),
                ReadBoolean(cursor, record, "fAutoFilter"),
                ReadBoolean(cursor, record, "fPivotTables"),
                ReadBoolean(cursor, record, "fSelUnlockedCells"));
        }

        private static bool ReadBoolean(XlsbBinaryCursor cursor, XlsbRecord record, string fieldName) =>
            XlsbBinaryValueReader.ReadUInt32Boolean(cursor, record, fieldName);
    }
}
