using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffInlineDifferentialFormatReader {
        private const int DxfNFlagsSize = 6;
        private const int DxfFntDSize = 118;
        private const int StxpOffset = 64;
        private const int DxfPatSize = 4;

        internal static bool TryReadDxfN(
            byte[] payload,
            int offset,
            int length,
            LegacyXlsWorkbook workbook,
            ushort recordType,
            int recordOffset,
            out LegacyXlsDifferentialFormat? differentialFormat) {
            differentialFormat = null;
            if (length < DxfNFlagsSize || offset < 0 || offset + length > payload.Length) {
                return false;
            }

            ulong flags = ReadUInt48(payload, offset);
            bool hasNumberFormat = IsBitSet(flags, 25);
            bool hasFont = IsBitSet(flags, 26);
            bool hasAlignment = IsBitSet(flags, 27);
            bool hasBorder = IsBitSet(flags, 28);
            bool hasPattern = IsBitSet(flags, 29);
            bool hasProtection = IsBitSet(flags, 30);
            if (hasNumberFormat
                || hasAlignment
                || hasBorder
                || hasProtection) {
                return false;
            }

            int endOffset = offset + length;
            int currentOffset = offset + DxfNFlagsSize;
            string? fontColor = null;
            bool? fontBold = null;
            bool? fontItalic = null;
            if (hasFont) {
                if (currentOffset + DxfFntDSize > endOffset) {
                    return false;
                }

                ReadFont(payload, currentOffset, workbook, out fontColor, out fontBold, out fontItalic);
                currentOffset += DxfFntDSize;
            }

            byte? fillPattern = null;
            string? foregroundColor = null;
            string? backgroundColor = null;
            if (hasPattern) {
                if (currentOffset + DxfPatSize > endOffset) {
                    return false;
                }

                ReadPattern(payload, currentOffset, flags, workbook, out fillPattern, out foregroundColor, out backgroundColor);
                currentOffset += DxfPatSize;
            }

            if (currentOffset != endOffset) {
                return false;
            }

            if (!fillPattern.HasValue
                && string.IsNullOrWhiteSpace(foregroundColor)
                && string.IsNullOrWhiteSpace(backgroundColor)
                && string.IsNullOrWhiteSpace(fontColor)
                && !fontBold.HasValue
                && !fontItalic.HasValue) {
                return false;
            }

            differentialFormat = new LegacyXlsDifferentialFormat(
                index: -1,
                fillPattern,
                foregroundColor,
                backgroundColor,
                fontColor,
                fontBold,
                fontItalic,
                recordType,
                recordOffset);
            return true;
        }

        private static void ReadPattern(
            byte[] payload,
            int patternOffset,
            ulong flags,
            LegacyXlsWorkbook workbook,
            out byte? fillPattern,
            out string? foregroundColor,
            out string? backgroundColor) {
            bool ignorePattern = IsBitSet(flags, 16);
            bool ignoreForeground = IsBitSet(flags, 17);
            bool ignoreBackground = IsBitSet(flags, 18);
            uint patternBits = BiffRecordReader.ReadUInt32(payload, patternOffset);
            fillPattern = ignorePattern ? null : (byte)((patternBits >> 10) & 0x3f);
            ushort foregroundIndex = (ushort)((patternBits >> 16) & 0x7f);
            ushort backgroundIndex = (ushort)((patternBits >> 23) & 0x7f);

            foregroundColor = null;
            backgroundColor = null;
            if (!ignoreForeground) {
                workbook.TryResolveColor(foregroundIndex, out foregroundColor);
            }

            if (!ignoreBackground) {
                workbook.TryResolveColor(backgroundIndex, out backgroundColor);
            }
        }

        private static void ReadFont(
            byte[] payload,
            int fontOffset,
            LegacyXlsWorkbook workbook,
            out string? fontColor,
            out bool? fontBold,
            out bool? fontItalic) {
            int stxpOffset = fontOffset + StxpOffset;
            uint textAttributes = BiffRecordReader.ReadUInt32(payload, stxpOffset + 4);
            ushort boldWeight = BiffRecordReader.ReadUInt16(payload, stxpOffset + 8);
            int colorIndex = BiffRecordReader.ReadInt32(payload, stxpOffset + 16);
            uint textAttributesNotChanged = BiffRecordReader.ReadUInt32(payload, stxpOffset + 24);
            uint boldWeightNotChanged = BiffRecordReader.ReadUInt32(payload, stxpOffset + 36);

            fontColor = null;
            if (colorIndex >= 0
                && colorIndex != 32767
                && workbook.TryResolveColor((ushort)colorIndex, out string? resolvedColor)) {
                fontColor = resolvedColor;
            }

            fontBold = null;
            if (boldWeightNotChanged == 0 && boldWeight != 0xffff) {
                fontBold = boldWeight >= 0x02bc;
            }

            fontItalic = null;
            if (!IsBitSet(textAttributesNotChanged, 1)) {
                fontItalic = IsBitSet(textAttributes, 1);
            }
        }

        private static ulong ReadUInt48(byte[] payload, int offset) {
            ulong value = 0;
            for (int i = 0; i < DxfNFlagsSize; i++) {
                value |= (ulong)payload[offset + i] << (8 * i);
            }

            return value;
        }

        private static bool IsBitSet(ulong value, int bit) {
            return ((value >> bit) & 1UL) != 0;
        }
    }
}
