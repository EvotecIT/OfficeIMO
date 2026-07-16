using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Decodes the common equality-list AutoFilter subset.</summary>
    internal static class XlsbWorksheetAutoFilterReader {
        internal static XlsbAutoFilterColumn ReadColumn(XlsbRecord record, XlsbAutoFilter autoFilter) {
            if (record.Data.Length != 6) {
                throw new InvalidDataException($"The BrtBeginFilterColumn record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }
            var cursor = new XlsbBinaryCursor(record.Data);
            uint columnId = cursor.ReadUInt32();
            ushort flags = cursor.ReadUInt16();
            uint width = checked((uint)(autoFilter.Range.LastColumn - autoFilter.Range.FirstColumn + 1));
            if (columnId >= width || autoFilter.Columns.Any(column => column.ColumnId == columnId) || (flags & 0xFFFC) != 0) {
                throw new InvalidDataException($"The BrtBeginFilterColumn record at offset {record.Offset} contains an invalid column or reserved flags.");
            }
            return new XlsbAutoFilterColumn(columnId) {
                HasUnsupportedContent = (flags & 0x0003) != 0
            };
        }

        internal static void ReadBeginFilters(XlsbRecord record, XlsbAutoFilterColumn column) {
            if (record.Data.Length != 8) {
                throw new InvalidDataException($"The BrtBeginFilters record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }
            var cursor = new XlsbBinaryCursor(record.Data);
            column.IncludeBlank = XlsbBinaryValueReader.ReadUInt32Boolean(cursor, record, "fBlank");
            cursor.ReadUInt32(); // unused
        }

        internal static string ReadValue(XlsbRecord record) {
            var cursor = new XlsbBinaryCursor(record.Data);
            string value = cursor.ReadWideString(255);
            if (value.Length == 0 || cursor.Remaining != 0) {
                throw new InvalidDataException($"The BrtFilter record at offset {record.Offset} contains an invalid value.");
            }
            return value;
        }
    }
}
