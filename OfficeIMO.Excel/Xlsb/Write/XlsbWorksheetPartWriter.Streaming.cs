using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Write {
    internal static partial class XlsbWorksheetPartWriter {
        /// <summary>Writes an already validated cell snapshot to the worksheet ZIP entry.</summary>
        internal static void WriteDirectTabular(
            Stream destination,
            XlsbDirectTabularPlan plan,
            CancellationToken cancellationToken) {
            using var writer = new XlsbDirectRecordWriter(destination);
            writer.WriteRecord(129); // BrtBeginSheet
            writer.WriteHeader(BrtWsDim, 16);
            WriteDirectDimension(writer, plan.RowCount, plan.ColumnCount);
            writer.WriteRecord(BrtBeginSheetData);

            for (int row = 0; row < plan.RowCount; row++) {
                if ((row & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (plan.ColumnCount != 0) WriteDirectRowHeader(writer, row, plan.ColumnCount);
                bool wroteCell = false;
                int rowStart = row * plan.ColumnCount;
                for (int column = 0; column < plan.ColumnCount; column++) {
                    int slot = rowStart + column;
                    byte kind = plan.KindAt(slot);
                    if (kind == (byte)ExcelDirectTabularValueKind.Empty) continue;
                    wroteCell = true;
                    switch (kind) {
                        case (byte)ExcelDirectTabularValueKind.Text:
                            if (plan.UsesSharedStrings) {
                                writer.WriteSharedStringCell(BrtCellIsst, column, checked((int)plan.PayloadAt(slot)));
                            } else {
                                writer.WriteTextCell(BrtCellSt, column, plan.TextAt(slot));
                            }
                            break;
                        case (byte)ExcelDirectTabularValueKind.Boolean:
                            WriteDirectBooleanCell(writer, column, plan.PayloadAt(slot) != 0);
                            break;
                        case (byte)ExcelDirectTabularValueKind.Number:
                        case XlsbDirectTabularPlan.DateKind:
                            writer.WriteNumberCell(BrtCellReal, column,
                                BitConverter.Int64BitsToDouble(unchecked((long)plan.PayloadAt(slot))),
                                kind == XlsbDirectTabularPlan.DateKind ? plan.DateStyle : 0U);
                            break;
                    }
                }
                if (!wroteCell && plan.ColumnCount != 0) {
                    writer.WriteHeader(BrtCellBlank, 8);
                    writer.WriteUInt32(checked((uint)(plan.ColumnCount - 1)));
                    writer.WriteUInt32(0U);
                }
            }

            writer.WriteRecord(BrtEndSheetData);
            writer.WriteRecord(BrtEndSheet);
            writer.Flush();
        }
    }
}
