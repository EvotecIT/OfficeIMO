using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsCellWatchWriter {
        private const ushort CellWatchRecordType = 0x086C;

        internal static bool SupportsWorksheetCellWatches(ExcelSheet sheet, out string? reason) {
            reason = null;
            CellWatches? cellWatches = GetWorksheetCellWatchesCollection(sheet);
            if (cellWatches == null) {
                return true;
            }

            if (!SupportsCellWatchesCollection(cellWatches)) {
                reason = "cell watches with unsupported metadata";
                return false;
            }

            foreach (CellWatch cellWatch in cellWatches.Elements<CellWatch>()) {
                if (!TryCreatePayload(cellWatch, out _, out reason)) {
                    return false;
                }
            }

            return true;
        }

        internal static IReadOnlyList<byte[]> CreateCellWatchPayloads(ExcelSheet sheet) {
            var payloads = new List<byte[]>();
            foreach (CellWatch cellWatch in GetWorksheetCellWatches(sheet)) {
                if (TryCreatePayload(cellWatch, out byte[]? payload, out _)) {
                    payloads.Add(payload!);
                }
            }

            return payloads;
        }

        private static IReadOnlyList<CellWatch> GetWorksheetCellWatches(ExcelSheet sheet) {
            CellWatches? cellWatches = GetWorksheetCellWatchesCollection(sheet);
            return cellWatches == null
                ? Array.Empty<CellWatch>()
                : cellWatches.Elements<CellWatch>().ToArray();
        }

        private static CellWatches? GetWorksheetCellWatchesCollection(ExcelSheet sheet) {
            return sheet.WorksheetPart.Worksheet?.Elements<CellWatches>().FirstOrDefault();
        }

        private static bool TryCreatePayload(CellWatch cellWatch, out byte[]? payload, out string? reason) {
            payload = null;
            reason = null;

            if (cellWatch.HasChildren || !SupportsKnownAttributes(cellWatch)) {
                reason = "cell watches with unsupported metadata";
                return false;
            }

            string? reference = cellWatch.CellReference?.Value;
            if (string.IsNullOrWhiteSpace(reference)) {
                reason = "cell watches without references";
                return false;
            }

            string normalized = reference!.Replace("$", string.Empty);
            if (normalized.Contains("!", StringComparison.Ordinal) || normalized.Contains(":", StringComparison.Ordinal)) {
                reason = "cell watch references";
                return false;
            }

            if (!A1.TryParseCellReferenceFast(normalized, out int row, out int column)) {
                reason = "cell watch references";
                return false;
            }

            if (row < 1 || row > 65536 || column < 1 || column > 256) {
                reason = "cell watch references outside BIFF8 worksheet limits";
                return false;
            }

            payload = BuildPayload(checked((ushort)(row - 1)), checked((ushort)(column - 1)));
            return true;
        }

        private static bool SupportsKnownAttributes(OpenXmlElement element) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri) || attribute.LocalName != "r") {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsCellWatchesCollection(CellWatches cellWatches) {
            if (cellWatches.ChildElements.Any(child => child is not CellWatch)) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in cellWatches.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                if (!string.Equals(attribute.LocalName, "count", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return true;
        }

        private static byte[] BuildPayload(ushort row, ushort column) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, CellWatchRecordType);
            WriteUInt16(stream, 0x0001);
            WriteUInt16(stream, row);
            WriteUInt16(stream, row);
            WriteUInt16(stream, column);
            WriteUInt16(stream, column);
            WriteUInt32(stream, 0);
            return stream.ToArray();
        }

        private static void WriteUInt16(Stream stream, int value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
            stream.WriteByte((byte)((value >> 16) & 0xff));
            stream.WriteByte((byte)((value >> 24) & 0xff));
        }
    }
}
