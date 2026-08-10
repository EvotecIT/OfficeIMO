using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private static void WriteWorkbookOptionRecords(Stream stream, ExcelDocument document) {
            WorkbookProperties? properties = document.WorkbookRoot.GetFirstChild<WorkbookProperties>();
            WriteRecord(stream, 0x0040, BuildUInt16Payload(
                properties?.BackupFile?.Value == true ? (ushort)1 : (ushort)0));
            WriteRecord(stream, 0x008d, BuildUInt16Payload(
                properties?.ShowObjects?.Value is ObjectDisplayValues showObjects
                    ? ToHiddenObjectsMode(showObjects)
                    : (ushort)0));
        }

        private static void WriteWorkbookPostCalculationOptionRecords(Stream stream, ExcelDocument document) {
            WorkbookProperties? properties = document.WorkbookRoot.GetFirstChild<WorkbookProperties>();
            WriteRecord(stream, 0x01b7, BuildUInt16Payload(
                properties?.RefreshAllConnections?.Value == true ? (ushort)1 : (ushort)0));
            WriteRecord(stream, 0x00da, BuildUInt16Payload(
                properties == null ? (ushort)0 : BuildBookBoolFlags(properties) ?? 0));
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
