using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffDifferentialFormatReader {
        private const int FrtHeaderSize = 12;
        private const int DxfFlagsSize = 2;
        private const int XfPropsHeaderSize = 4;
        private const int XfPropHeaderSize = 4;
        private const ushort DxfRecordType = (ushort)BiffRecordType.Dxf;

        internal static bool TryRead(
            BiffRecord record,
            LegacyXlsWorkbook workbook,
            out LegacyXlsDifferentialFormat? differentialFormat) {
            differentialFormat = null;
            byte[] payload = record.Payload;
            int xfPropsOffset = FrtHeaderSize + DxfFlagsSize;
            if (payload.Length < xfPropsOffset + XfPropsHeaderSize
                || BiffRecordReader.ReadUInt16(payload, 0) != DxfRecordType) {
                return false;
            }

            int offset = xfPropsOffset;
            ushort propertyCount = BiffRecordReader.ReadUInt16(payload, offset + 2);
            offset += XfPropsHeaderSize;
            if (propertyCount > 1024) {
                return false;
            }

            byte? fillPattern = null;
            string? fillForegroundColor = null;
            string? fillBackgroundColor = null;

            for (int i = 0; i < propertyCount; i++) {
                if (offset + XfPropHeaderSize > payload.Length) {
                    return false;
                }

                ushort propertyType = BiffRecordReader.ReadUInt16(payload, offset);
                ushort propertySize = BiffRecordReader.ReadUInt16(payload, offset + 2);
                if (propertySize < XfPropHeaderSize || offset + propertySize > payload.Length) {
                    return false;
                }

                int dataOffset = offset + XfPropHeaderSize;
                int dataLength = propertySize - XfPropHeaderSize;
                if (propertyType == 0x0000 && dataLength >= 1) {
                    fillPattern = payload[dataOffset];
                } else if (propertyType == 0x0001 && TryReadColor(payload, dataOffset, dataLength, workbook, out string? foregroundColor)) {
                    fillForegroundColor = foregroundColor;
                } else if (propertyType == 0x0002 && TryReadColor(payload, dataOffset, dataLength, workbook, out string? backgroundColor)) {
                    fillBackgroundColor = backgroundColor;
                }

                offset += propertySize;
            }

            if (!fillPattern.HasValue && fillForegroundColor == null && fillBackgroundColor == null) {
                return false;
            }

            differentialFormat = new LegacyXlsDifferentialFormat(
                workbook.DifferentialFormats.Count,
                fillPattern,
                fillForegroundColor,
                fillBackgroundColor,
                record.Type,
                record.Offset);
            return true;
        }

        private static bool TryReadColor(
            byte[] payload,
            int offset,
            int length,
            LegacyXlsWorkbook workbook,
            out string? argb) {
            argb = null;
            if (length < 8 || offset + 8 > payload.Length) {
                return false;
            }

            byte flags = payload[offset];
            if ((flags & 0x01) == 0) {
                return false;
            }

            byte colorType = (byte)(flags >> 1);
            byte indexedColor = payload[offset + 1];
            if (colorType == 0x01) {
                return workbook.TryResolveColor(indexedColor, out argb);
            }

            if (colorType == 0x02) {
                byte red = payload[offset + 4];
                byte green = payload[offset + 5];
                byte blue = payload[offset + 6];
                byte alpha = payload[offset + 7];
                argb = alpha.ToString("X2") + red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");
                return true;
            }

            return false;
        }
    }
}
