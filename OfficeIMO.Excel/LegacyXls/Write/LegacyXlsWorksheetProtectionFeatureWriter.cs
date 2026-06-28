using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsWorksheetProtectionFeatureWriter {
        private const ushort FeatHdrRecordType = 0x0867;
        private const ushort IsfProtection = 0x0002;

        internal static bool TryCreatePayload(SheetProtection? protection, out byte[]? payload) {
            payload = null;
            if (protection == null || protection.Sheet?.Value == false) {
                return false;
            }

            LegacyXlsWorksheetProtectionPermissions permissions = ExtractPermissions(protection);
            if (permissions.Equals(LegacyXlsWorksheetProtectionPermissions.Default(protection.Objects?.Value, protection.Scenarios?.Value))) {
                return false;
            }

            payload = BuildPayload(permissions);
            return true;
        }

        internal static LegacyXlsWorksheetProtectionPermissions ExtractPermissions(SheetProtection protection) {
            if (protection == null) throw new ArgumentNullException(nameof(protection));

            return new LegacyXlsWorksheetProtectionPermissions(
                allowEditObjects: !(protection.Objects?.Value ?? false),
                allowEditScenarios: !(protection.Scenarios?.Value ?? false),
                allowFormatCells: IsProtectionActionAllowed(protection.FormatCells, lockedWhenOmitted: true),
                allowFormatColumns: IsProtectionActionAllowed(protection.FormatColumns, lockedWhenOmitted: true),
                allowFormatRows: IsProtectionActionAllowed(protection.FormatRows, lockedWhenOmitted: true),
                allowInsertColumns: IsProtectionActionAllowed(protection.InsertColumns, lockedWhenOmitted: true),
                allowInsertRows: IsProtectionActionAllowed(protection.InsertRows, lockedWhenOmitted: true),
                allowInsertHyperlinks: IsProtectionActionAllowed(protection.InsertHyperlinks, lockedWhenOmitted: true),
                allowDeleteColumns: IsProtectionActionAllowed(protection.DeleteColumns, lockedWhenOmitted: true),
                allowDeleteRows: IsProtectionActionAllowed(protection.DeleteRows, lockedWhenOmitted: true),
                allowSelectLockedCells: IsProtectionActionAllowed(protection.SelectLockedCells, lockedWhenOmitted: false),
                allowSort: IsProtectionActionAllowed(protection.Sort, lockedWhenOmitted: true),
                allowAutoFilter: IsProtectionActionAllowed(protection.AutoFilter, lockedWhenOmitted: true),
                allowPivotTables: IsProtectionActionAllowed(protection.PivotTables, lockedWhenOmitted: true),
                allowSelectUnlockedCells: IsProtectionActionAllowed(protection.SelectUnlockedCells, lockedWhenOmitted: false));
        }

        private static byte[] BuildPayload(LegacyXlsWorksheetProtectionPermissions permissions) {
            byte[] payload = new byte[23];
            WriteUInt16(payload, 0, FeatHdrRecordType);
            WriteUInt16(payload, 2, 0);
            WriteUInt16(payload, 12, IsfProtection);
            payload[14] = 1;
            WriteUInt32(payload, 15, 0xffffffff);
            WriteUInt32(payload, 19, BuildEnhancedProtectionFlags(permissions));
            return payload;
        }

        private static uint BuildEnhancedProtectionFlags(LegacyXlsWorksheetProtectionPermissions permissions) {
            uint flags = 0;
            SetBit(ref flags, 2, permissions.AllowEditObjects);
            SetBit(ref flags, 3, permissions.AllowEditScenarios);
            SetBit(ref flags, 4, permissions.AllowFormatCells);
            SetBit(ref flags, 5, permissions.AllowFormatColumns);
            SetBit(ref flags, 6, permissions.AllowFormatRows);
            SetBit(ref flags, 7, permissions.AllowInsertColumns);
            SetBit(ref flags, 8, permissions.AllowInsertRows);
            SetBit(ref flags, 9, permissions.AllowInsertHyperlinks);
            SetBit(ref flags, 10, permissions.AllowDeleteColumns);
            SetBit(ref flags, 11, permissions.AllowDeleteRows);
            SetBit(ref flags, 12, permissions.AllowSelectLockedCells);
            SetBit(ref flags, 13, permissions.AllowSort);
            SetBit(ref flags, 14, permissions.AllowAutoFilter);
            SetBit(ref flags, 15, permissions.AllowPivotTables);
            SetBit(ref flags, 16, permissions.AllowSelectUnlockedCells);
            return flags;
        }

        private static bool IsProtectionActionAllowed(BooleanValue? lockFlag, bool lockedWhenOmitted) {
            return !(lockFlag?.Value ?? lockedWhenOmitted);
        }

        private static void SetBit(ref uint flags, int bit, bool value) {
            if (value) {
                flags |= 1u << bit;
            }
        }

        private static void WriteUInt16(byte[] buffer, int offset, ushort value) {
            buffer[offset] = (byte)(value & 0xff);
            buffer[offset + 1] = (byte)((value >> 8) & 0xff);
        }

        private static void WriteUInt32(byte[] buffer, int offset, uint value) {
            buffer[offset] = (byte)(value & 0xff);
            buffer[offset + 1] = (byte)((value >> 8) & 0xff);
            buffer[offset + 2] = (byte)((value >> 16) & 0xff);
            buffer[offset + 3] = (byte)((value >> 24) & 0xff);
        }
    }
}
