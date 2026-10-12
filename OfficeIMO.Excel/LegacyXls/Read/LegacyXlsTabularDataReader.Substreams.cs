using OfficeIMO.Excel.LegacyXls.Biff;
using static OfficeIMO.Excel.LegacyXls.Read.LegacyXlsTabularWorkbook;

namespace OfficeIMO.Excel.LegacyXls.Read {
    internal sealed partial class LegacyXlsTabularDataReader {
        // Embedded charts have their own BOF/EOF pair and may contain cell-shaped
        // cached series records. Those records do not belong to the worksheet's rows.
        private void SkipNestedChartSubstream(RecordSlice bof, ref int position) {
            ValidateNestedChartBof(bof);
            int depth = 1;
            while (position < _worksheetLimitOffset
                   && TryReadDiscoveryRecord(ref position, out RecordSlice record)) {
                if (CanCancelCurrentRead) CheckCancellation();
                if (record.PayloadOffset + record.Length > _worksheetLimitOffset) {
                    throw new InvalidDataException("An embedded XLS chart record extends beyond its worksheet boundary.");
                }
                if (record.Type == (ushort)BiffRecordType.Bof) {
                    ValidateNestedChartBof(record);
                    depth++;
                } else if (record.Type == (ushort)BiffRecordType.Eof && --depth == 0) {
                    return;
                }
            }

            throw new InvalidDataException("The embedded XLS chart substream is truncated before EOF.");
        }

        private void ValidateNestedChartBof(RecordSlice bof) {
            if (bof.Length < 4
                || ReadDiscoveryUInt16(bof.PayloadOffset) != 0x0600
                || ReadDiscoveryUInt16(bof.PayloadOffset + 2) != 0x0020) {
                throw new InvalidDataException("An embedded XLS substream must start with a valid BIFF8 chart BOF record.");
            }
        }
    }
}
