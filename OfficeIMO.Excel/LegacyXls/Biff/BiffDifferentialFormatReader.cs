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
            string? fontColor = null;
            bool? fontBold = null;
            bool? fontItalic = null;
            LegacyXlsDifferentialBorderSide? topBorder = null;
            LegacyXlsDifferentialBorderSide? bottomBorder = null;
            LegacyXlsDifferentialBorderSide? leftBorder = null;
            LegacyXlsDifferentialBorderSide? rightBorder = null;
            ushort? numberFormatId = null;
            string? numberFormatCode = null;

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
                } else if (propertyType == 0x0005 && TryReadColor(payload, dataOffset, dataLength, workbook, out string? textColor)) {
                    fontColor = textColor;
                } else if (propertyType == 0x0019 && dataLength >= 2) {
                    ushort weight = BiffRecordReader.ReadUInt16(payload, dataOffset);
                    fontBold = weight >= 0x02bc;
                } else if (propertyType == 0x001c && dataLength >= 1) {
                    fontItalic = payload[dataOffset] != 0;
                } else if (propertyType == 0x0026
                    && TryReadNumberFormatCode(payload, dataOffset, dataLength, out string? inlineNumberFormatCode)) {
                    numberFormatCode = inlineNumberFormatCode;
                } else if (propertyType == 0x0029 && dataLength >= 2) {
                    numberFormatId = BiffRecordReader.ReadUInt16(payload, dataOffset);
                    if (TryResolveNumberFormatCode(workbook, numberFormatId.Value, out string? resolvedNumberFormatCode)) {
                        numberFormatCode = resolvedNumberFormatCode;
                    }
                } else if (propertyType >= 0x0006
                    && propertyType <= 0x0009
                    && TryReadBorderSide(payload, dataOffset, dataLength, workbook, out LegacyXlsDifferentialBorderSide? borderSide)) {
                    switch (propertyType) {
                        case 0x0006:
                            topBorder = borderSide;
                            break;
                        case 0x0007:
                            bottomBorder = borderSide;
                            break;
                        case 0x0008:
                            leftBorder = borderSide;
                            break;
                        case 0x0009:
                            rightBorder = borderSide;
                            break;
                    }
                }

                offset += propertySize;
            }

            LegacyXlsDifferentialBorder? border = CreateBorder(topBorder, bottomBorder, leftBorder, rightBorder);
            if (!fillPattern.HasValue
                && fillForegroundColor == null
                && fillBackgroundColor == null
                && fontColor == null
                && !fontBold.HasValue
                && !fontItalic.HasValue
                && border == null
                && !numberFormatId.HasValue
                && string.IsNullOrWhiteSpace(numberFormatCode)) {
                return false;
            }

            differentialFormat = new LegacyXlsDifferentialFormat(
                workbook.DifferentialFormats.Count,
                fillPattern,
                fillForegroundColor,
                fillBackgroundColor,
                fontColor,
                fontBold,
                fontItalic,
                record.Type,
                record.Offset,
                border,
                numberFormatId,
                numberFormatCode);
            return true;
        }

        private static LegacyXlsDifferentialBorder? CreateBorder(
            LegacyXlsDifferentialBorderSide? top,
            LegacyXlsDifferentialBorderSide? bottom,
            LegacyXlsDifferentialBorderSide? left,
            LegacyXlsDifferentialBorderSide? right) {
            var border = new LegacyXlsDifferentialBorder(top, bottom, left, right);
            return border.HasAnySide ? border : null;
        }

        private static bool TryReadBorderSide(
            byte[] payload,
            int offset,
            int length,
            LegacyXlsWorkbook workbook,
            out LegacyXlsDifferentialBorderSide? borderSide) {
            borderSide = null;
            if (length < 10 || offset + 10 > payload.Length) {
                return false;
            }

            string? color = null;
            _ = TryReadColor(payload, offset, 8, workbook, out color);
            ushort style = BiffRecordReader.ReadUInt16(payload, offset + 8);
            if (style == 0 && string.IsNullOrWhiteSpace(color)) {
                return false;
            }

            borderSide = new LegacyXlsDifferentialBorderSide(style, color);
            return true;
        }

        private static bool TryReadNumberFormatCode(byte[] payload, int offset, int length, out string? numberFormatCode) {
            numberFormatCode = null;
            if (length <= 0 || offset + length > payload.Length) {
                return false;
            }

            try {
                byte[] formatPayload = new byte[length];
                Array.Copy(payload, offset, formatPayload, 0, length);
                int stringOffset = 0;
                string value = BiffStringReader.ReadUnicodeString(formatPayload, ref stringOffset);
                if (stringOffset > formatPayload.Length || string.IsNullOrWhiteSpace(value)) {
                    return false;
                }

                numberFormatCode = value;
                return true;
            } catch (Exception ex) when (ex is InvalidDataException || ex is OverflowException) {
                return false;
            }
        }

        private static bool TryResolveNumberFormatCode(LegacyXlsWorkbook workbook, ushort numberFormatId, out string? numberFormatCode) {
            if (BiffBuiltInNumberFormat.TryGetCode(numberFormatId, out numberFormatCode)) {
                return true;
            }

            LegacyXlsNumberFormat? format = workbook.NumberFormats.FirstOrDefault(format => format.FormatId == numberFormatId);
            numberFormatCode = format?.FormatCode;
            return !string.IsNullOrWhiteSpace(numberFormatCode);
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
