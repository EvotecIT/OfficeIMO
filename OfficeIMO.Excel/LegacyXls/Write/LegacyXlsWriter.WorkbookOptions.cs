using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private static void WriteWorkbookOptionRecords(Stream stream, ExcelDocument document) {
            WorkbookProperties? properties = document.WorkbookRoot.GetFirstChild<WorkbookProperties>();
            if (properties == null) {
                return;
            }

            if (properties.BackupFile?.Value is bool backupFile) {
                WriteRecord(stream, 0x0040, BuildUInt16Payload(backupFile ? (ushort)1 : (ushort)0));
            }

            ushort? bookBoolFlags = BuildBookBoolFlags(properties);
            if (bookBoolFlags.HasValue) {
                WriteRecord(stream, 0x00da, BuildUInt16Payload(bookBoolFlags.Value));
            }

            if (properties.ShowObjects?.Value is ObjectDisplayValues showObjects) {
                WriteRecord(stream, 0x008d, BuildUInt16Payload(ToHiddenObjectsMode(showObjects)));
            }

            if (properties.RefreshAllConnections?.Value == true) {
                WriteRecord(stream, 0x01b7, Array.Empty<byte>());
            }
        }

        private static ushort? BuildBookBoolFlags(WorkbookProperties properties) {
            bool hasMappedFlag = false;
            ushort flags = 0;

            if (properties.SaveExternalLinkValues?.Value is bool saveExternalLinkValues) {
                hasMappedFlag = true;
                if (!saveExternalLinkValues) {
                    flags |= 0x0001;
                }
            }

            if (properties.ShowBorderUnselectedTables?.Value is bool showBorderUnselectedTables) {
                hasMappedFlag = true;
                if (!showBorderUnselectedTables) {
                    flags |= 0x0100;
                }
            }

            return hasMappedFlag ? flags : null;
        }

        private static ushort ToHiddenObjectsMode(ObjectDisplayValues showObjects) {
            if (showObjects == ObjectDisplayValues.None) {
                return 2;
            }

            if (showObjects == ObjectDisplayValues.Placeholders) {
                return 1;
            }

            return 0;
        }
    }
}
