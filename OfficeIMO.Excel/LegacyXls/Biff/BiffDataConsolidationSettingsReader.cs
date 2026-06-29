using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffDataConsolidationSettingsReader {
        internal static bool TryRead(BiffRecord record, out LegacyXlsDataConsolidationSettings? settings) {
            settings = null;
            if (record.Type != (ushort)BiffRecordType.DCon || record.Payload.Length < 4) {
                return false;
            }

            ushort rawFunction = BiffRecordReader.ReadUInt16(record.Payload, 0);
            LegacyXlsDataConsolidationFunction? function = TryGetFunction(rawFunction);
            if (!function.HasValue) {
                return false;
            }

            bool usesLeftLabels;
            bool usesTopLabels;
            bool linksToSourceData;
            if (record.Payload.Length >= 8) {
                usesLeftLabels = BiffRecordReader.ReadUInt16(record.Payload, 2) != 0;
                usesTopLabels = BiffRecordReader.ReadUInt16(record.Payload, 4) != 0;
                linksToSourceData = BiffRecordReader.ReadUInt16(record.Payload, 6) != 0;
            } else {
                ushort legacyFlags = BiffRecordReader.ReadUInt16(record.Payload, 2);
                usesTopLabels = (legacyFlags & 0x0001) != 0;
                usesLeftLabels = (legacyFlags & 0x0002) != 0;
                linksToSourceData = (legacyFlags & 0x0004) != 0;
            }

            ushort flags = CreateOptionFlags(usesTopLabels, usesLeftLabels, linksToSourceData);
            settings = new LegacyXlsDataConsolidationSettings(
                function.Value,
                usesTopLabels,
                usesLeftLabels,
                linksToSourceData,
                rawFunction,
                flags);
            return true;
        }

        private static ushort CreateOptionFlags(bool usesTopLabels, bool usesLeftLabels, bool linksToSourceData) {
            ushort flags = 0;
            if (usesTopLabels) flags |= 0x0001;
            if (usesLeftLabels) flags |= 0x0002;
            if (linksToSourceData) flags |= 0x0004;
            return flags;
        }

        private static LegacyXlsDataConsolidationFunction? TryGetFunction(ushort value) {
            return value <= (ushort)LegacyXlsDataConsolidationFunction.VarianceP
                ? (LegacyXlsDataConsolidationFunction)value
                : null;
        }
    }
}
