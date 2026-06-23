using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffAutoFilterReader {
        internal static bool TryReadInfo(byte[] payload, out ushort dropDownCount) {
            dropDownCount = 0;
            if (payload.Length < 2) {
                return false;
            }

            dropDownCount = BiffRecordReader.ReadUInt16(payload, 0);
            return dropDownCount <= 256;
        }

        internal static bool TryReadCriteria(byte[] payload, out LegacyXlsAutoFilterCriteria? criteria) {
            criteria = null;
            if (payload.Length < 24) {
                return false;
            }

            ushort columnId = BiffRecordReader.ReadUInt16(payload, 0);
            if (columnId > 255) {
                return false;
            }

            ushort flags = BiffRecordReader.ReadUInt16(payload, 2);
            if ((flags & 0x0010) != 0) {
                ushort top10Value = (ushort)((flags >> 7) & 0x01ff);
                if (top10Value == 0 || top10Value > 500) {
                    return false;
                }

                bool isTop = (flags & 0x0020) != 0;
                bool isPercent = (flags & 0x0040) != 0;
                criteria = LegacyXlsAutoFilterCriteria.CreateTop10(columnId, top10Value, isTop, isPercent);
                return true;
            }

            int stringOffset = 24;
            var conditions = new List<LegacyXlsAutoFilterCondition>(2);
            LegacyXlsAutoFilterKind? firstKind;
            LegacyXlsAutoFilterKind? secondKind;
            if (!TryReadDoper(payload, 4, ref stringOffset, out LegacyXlsAutoFilterCondition? firstCondition, out firstKind)
                || !TryReadDoper(payload, 14, ref stringOffset, out LegacyXlsAutoFilterCondition? secondCondition, out secondKind)) {
                return false;
            }

            if (firstCondition != null) {
                conditions.Add(firstCondition);
            }

            if (secondCondition != null) {
                conditions.Add(secondCondition);
            }

            if (conditions.Count == 0) {
                return false;
            }

            bool matchAll = (flags & 0x0003) == 0x0001;
            LegacyXlsAutoFilterKind kind = ResolveCriteriaKind(conditions, firstKind, secondKind);
            criteria = new LegacyXlsAutoFilterCriteria(columnId, matchAll, conditions, kind);
            return true;
        }

        private static bool TryReadDoper(
            byte[] payload,
            int offset,
            ref int stringOffset,
            out LegacyXlsAutoFilterCondition? condition,
            out LegacyXlsAutoFilterKind? specialKind) {
            condition = null;
            specialKind = null;
            if (offset + 10 > payload.Length) {
                return false;
            }

            byte valueType = payload[offset];
            if (valueType == 0x00) {
                return true;
            }

            if (!TryMapOperator(payload[offset + 1], out LegacyXlsAutoFilterOperator @operator)) {
                return false;
            }

            string value;
            switch (valueType) {
                case 0x02:
                    value = BiffRkNumberReader.ReadRkNumber(BiffRecordReader.ReadUInt32(payload, offset + 2)).ToString("G15", CultureInfo.InvariantCulture);
                    break;
                case 0x04:
                    value = BiffRecordReader.ReadDouble(payload, offset + 2).ToString("G15", CultureInfo.InvariantCulture);
                    break;
                case 0x06:
                    int charCount = payload[offset + 6];
                    value = BiffStringReader.ReadUnicodeStringNoCch(payload, ref stringOffset, charCount);
                    break;
                case 0x08:
                    if (payload[offset + 2] != 0) {
                        value = BiffErrorValue.ToText(payload[offset + 3]);
                    } else {
                        value = payload[offset + 3] == 0 ? "FALSE" : "TRUE";
                    }

                    break;
                case 0x0c:
                    if (@operator != LegacyXlsAutoFilterOperator.Equal) {
                        return false;
                    }

                    value = string.Empty;
                    specialKind = LegacyXlsAutoFilterKind.Blanks;
                    break;
                case 0x0e:
                    if (@operator != LegacyXlsAutoFilterOperator.NotEqual) {
                        return false;
                    }

                    value = string.Empty;
                    specialKind = LegacyXlsAutoFilterKind.NonBlanks;
                    break;
                default:
                    return false;
            }

            condition = new LegacyXlsAutoFilterCondition(@operator, value);
            return true;
        }

        private static LegacyXlsAutoFilterKind ResolveCriteriaKind(
            IReadOnlyList<LegacyXlsAutoFilterCondition> conditions,
            LegacyXlsAutoFilterKind? firstKind,
            LegacyXlsAutoFilterKind? secondKind) {
            if (conditions.Count == 1) {
                return firstKind ?? secondKind ?? LegacyXlsAutoFilterKind.Custom;
            }

            return LegacyXlsAutoFilterKind.Custom;
        }

        private static bool TryMapOperator(byte value, out LegacyXlsAutoFilterOperator @operator) {
            switch (value) {
                case 0x01:
                    @operator = LegacyXlsAutoFilterOperator.LessThan;
                    return true;
                case 0x02:
                    @operator = LegacyXlsAutoFilterOperator.Equal;
                    return true;
                case 0x03:
                    @operator = LegacyXlsAutoFilterOperator.LessThanOrEqual;
                    return true;
                case 0x04:
                    @operator = LegacyXlsAutoFilterOperator.GreaterThan;
                    return true;
                case 0x05:
                    @operator = LegacyXlsAutoFilterOperator.NotEqual;
                    return true;
                case 0x06:
                    @operator = LegacyXlsAutoFilterOperator.GreaterThanOrEqual;
                    return true;
                default:
                    @operator = LegacyXlsAutoFilterOperator.Equal;
                    return false;
            }
        }
    }
}
