using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal static class XlsbWorksheetViewReader {
        internal static bool ReadGridlines(XlsbRecord record) {
            if (record.Data.Length != 30) {
                throw new InvalidDataException($"The BrtBeginWsView record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }
            // MS-XLSB 2.4.307: fDspGrid is bit 2 in the first USHORT.
            return (new XlsbBinaryCursor(record.Data).ReadUInt16() & 0x0004) != 0;
        }
    }
}
