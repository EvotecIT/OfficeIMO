using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private static void WriteWorkbookCalculationSettingsRecords(Stream stream, ExcelDocument document) {
            CalculationProperties? properties = document.WorkbookRoot.GetFirstChild<CalculationProperties>();
            WriteRecord(stream, 0x000e, BuildUInt16Payload(
                properties?.FullPrecision?.Value == false ? (ushort)0 : (ushort)1));
        }

        private static void WriteWorksheetCalculationSettingsRecords(Stream stream, ExcelDocument document) {
            CalculationProperties? properties = document.WorkbookRoot.GetFirstChild<CalculationProperties>();

            WriteRecord(stream, 0x000d, BuildInt16Payload(
                properties?.CalculationMode?.Value is CalculateModeValues mode
                    ? ToBiffCalculationMode(mode)
                    : (short)1));

            uint iterationCount = properties?.IterateCount?.Value ?? 100U;
            if (iterationCount > short.MaxValue) {
                throw new NotSupportedException("Native XLS saving supports calculation iteration counts up to 32,767.");
            }
            WriteRecord(stream, 0x000c, BuildInt16Payload(checked((short)iterationCount)));

            WriteRecord(stream, 0x000f, BuildUInt16Payload(
                properties?.ReferenceMode?.Value == ReferenceModeValues.R1C1 ? (ushort)0 : (ushort)1));

            WriteRecord(stream, 0x0011, BuildUInt16Payload(
                properties?.Iterate?.Value == true ? (ushort)1 : (ushort)0));

            double iterateDelta = properties?.IterateDelta?.Value ?? 0.001d;
            if (double.IsNaN(iterateDelta) || double.IsInfinity(iterateDelta) || iterateDelta < 0d) {
                throw new NotSupportedException("Native XLS saving requires a non-negative finite calculation iteration delta.");
            }
            WriteRecord(stream, 0x0010, BuildDoublePayload(iterateDelta));

            WriteRecord(stream, 0x005f, BuildUInt16Payload(
                properties?.CalculationOnSave?.Value == false ? (ushort)0 : (ushort)1));
        }

        private static void WriteWorksheetCalculationRecords(Stream stream, ExcelSheet sheet) {
            SheetCalculationProperties? properties = sheet.WorksheetPart.Worksheet?.GetFirstChild<SheetCalculationProperties>();
            if (properties?.FullCalculationOnLoad?.Value == true) {
                WriteRecord(stream, 0x005e, BuildUInt16Payload(0));
            }
        }

        private static short ToBiffCalculationMode(CalculateModeValues mode) {
            if (mode == CalculateModeValues.Manual) {
                return 0;
            }

            if (mode == CalculateModeValues.AutoNoTable) {
                return 2;
            }

            return 1;
        }

        private static byte[] BuildInt16Payload(short value) {
            using var stream = new MemoryStream();
            WriteInt16(stream, value);
            return stream.ToArray();
        }
    }
}
