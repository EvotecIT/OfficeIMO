using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal static class XlsbWorksheetViewReader {
        private const int MaximumSelectionRanges = 8192;

        internal static bool ReadGridlines(XlsbRecord record) {
            if (record.Data.Length != 30) {
                throw new InvalidDataException($"The BrtBeginWsView record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }
            // MS-XLSB 2.4.307: fDspGrid is bit 2 in the first USHORT.
            return (new XlsbBinaryCursor(record.Data).ReadUInt16() & 0x0004) != 0;
        }

        /// <summary>Reads BrtBeginWsView (MS-XLSB 2.4.307); returns null for another window's view or out-of-range settings.</summary>
        internal static XlsbSheetView? ReadPrimaryView(XlsbRecord record) {
            if (record.Data.Length != 30) {
                throw new InvalidDataException($"The BrtBeginWsView record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }
            var cursor = new XlsbBinaryCursor(record.Data);
            ushort flags = cursor.ReadUInt16();
            uint view = cursor.ReadUInt32();
            uint topRow = cursor.ReadUInt32();
            uint leftColumn = cursor.ReadUInt32();
            byte gridColor = cursor.ReadByte();
            cursor.Skip(3);
            ushort zoom = cursor.ReadUInt16();
            ushort zoomNormal = cursor.ReadUInt16();
            ushort zoomPageBreak = cursor.ReadUInt16();
            ushort zoomPageLayout = cursor.ReadUInt16();
            uint workbookView = cursor.ReadUInt32();
            // Settings outside the documented ranges keep the gridline-only projection rather than failing the load.
            if (workbookView != 0U || view > 2U || topRow >= ExcelA1.MaxRows || leftColumn >= ExcelA1.MaxColumns
                || !IsZoom(zoom, allowZero: false) || !IsZoom(zoomNormal, allowZero: true)
                || !IsZoom(zoomPageBreak, allowZero: true) || !IsZoom(zoomPageLayout, allowZero: true)) {
                return null;
            }
            return new XlsbSheetView((ushort)(flags & 0x07FF), view, checked((int)topRow), checked((int)leftColumn), gridColor,
                zoom, zoomNormal, zoomPageBreak, zoomPageLayout);
        }

        /// <summary>Reads BrtSel (MS-XLSB 2.4.792): pnn, rwAct, colAct, dwRfxAct and an UncheckedSqRfX range list.</summary>
        internal static bool TryReadSelection(XlsbRecord record, out XlsbSelection? selection) {
            selection = null;
            byte[] data = record.Data;
            if (data.Length < 20) return false;
            var cursor = new XlsbBinaryCursor(data);
            uint pane = cursor.ReadUInt32();
            uint activeRow = cursor.ReadUInt32();
            uint activeColumn = cursor.ReadUInt32();
            uint activeCellId = cursor.ReadUInt32();
            uint count = cursor.ReadUInt32();
            if (pane > 3U || activeRow >= ExcelA1.MaxRows || activeColumn >= ExcelA1.MaxColumns
                || count == 0U || count > MaximumSelectionRanges || cursor.Remaining != count * 16U || activeCellId >= count) return false;
            var ranges = new XlsbCellRange[count];
            for (int index = 0; index < ranges.Length; index++) {
                uint firstRow = cursor.ReadUInt32();
                uint lastRow = cursor.ReadUInt32();
                uint firstColumn = cursor.ReadUInt32();
                uint lastColumn = cursor.ReadUInt32();
                if (firstRow > lastRow || lastRow >= ExcelA1.MaxRows || firstColumn > lastColumn || lastColumn >= ExcelA1.MaxColumns) return false;
                ranges[index] = new XlsbCellRange(checked((int)firstRow + 1), checked((int)lastRow + 1), checked((int)firstColumn + 1), checked((int)lastColumn + 1));
            }
            selection = new XlsbSelection(pane, checked((int)activeRow + 1), checked((int)activeColumn + 1), activeCellId, ranges);
            return true;
        }

        private static bool IsZoom(ushort value, bool allowZero) => (allowZero && value == 0) || (value >= 10 && value <= 400);
    }
}
