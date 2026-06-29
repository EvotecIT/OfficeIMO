using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private static void WriteCalculationSettingsRecords(Stream stream, ExcelDocument document) {
            CalculationProperties? properties = document.WorkbookRoot.GetFirstChild<CalculationProperties>();
            if (properties == null) {
                return;
            }

            if (properties.CalculationMode?.Value is CalculateModeValues mode) {
                WriteRecord(stream, 0x000d, BuildInt16Payload(ToBiffCalculationMode(mode)));
            }

            if (properties.IterateCount?.Value is uint iterationCount) {
                if (iterationCount > short.MaxValue) {
                    throw new NotSupportedException("Native XLS saving supports calculation iteration counts up to 32,767.");
                }

                WriteRecord(stream, 0x000c, BuildInt16Payload(checked((short)iterationCount)));
            }

            if (properties.FullPrecision?.Value is bool fullPrecision) {
                WriteRecord(stream, 0x000e, BuildUInt16Payload(fullPrecision ? (ushort)1 : (ushort)0));
            }

            if (properties.ReferenceMode?.Value is ReferenceModeValues referenceMode) {
                WriteRecord(stream, 0x000f, BuildUInt16Payload(referenceMode == ReferenceModeValues.R1C1 ? (ushort)0 : (ushort)1));
            }

            if (properties.IterateDelta?.Value is double iterateDelta) {
                if (double.IsNaN(iterateDelta) || double.IsInfinity(iterateDelta) || iterateDelta < 0d) {
                    throw new NotSupportedException("Native XLS saving requires a non-negative finite calculation iteration delta.");
                }

                WriteRecord(stream, 0x0010, BuildDoublePayload(iterateDelta));
            }

            if (properties.Iterate?.Value is bool iterate) {
                WriteRecord(stream, 0x0011, BuildUInt16Payload(iterate ? (ushort)1 : (ushort)0));
            }

            if (properties.CalculationOnSave?.Value is bool calculationOnSave) {
                WriteRecord(stream, 0x005f, BuildUInt16Payload(calculationOnSave ? (ushort)1 : (ushort)0));
            }
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
