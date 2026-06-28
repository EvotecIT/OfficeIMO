namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffPageLayoutViewReader {
        private const int FutureRecordHeaderLength = 12;
        private const int FixedPayloadLength = FutureRecordHeaderLength + 4;

        internal static bool TryRead(BiffRecord record, out uint? zoomScale) {
            zoomScale = null;
            if (record.Type != (ushort)BiffRecordType.Plv || record.Payload.Length < FutureRecordHeaderLength) {
                return false;
            }

            ushort wrappedRecordType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            if (wrappedRecordType != (ushort)BiffRecordType.Plv) {
                return false;
            }

            if (record.Payload.Length >= FixedPayloadLength) {
                ushort scale = BiffRecordReader.ReadUInt16(record.Payload, FutureRecordHeaderLength);
                if (scale >= 10 && scale <= 400) {
                    zoomScale = scale;
                }
            }

            return true;
        }
    }
}
