using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsIgnoredErrorWriter {
        private const ushort FeatHdrRecordType = 0x0867;
        private const ushort FeatRecordType = 0x0868;
        private const ushort IsfIgnoredErrors = 0x0003;

        internal static bool SupportsWorksheetIgnoredErrors(ExcelSheet sheet, out string? reason) {
            reason = null;
            IgnoredErrors? ignoredErrors = GetWorksheetIgnoredErrorsCollection(sheet);
            if (ignoredErrors == null) {
                return true;
            }

            if (!SupportsIgnoredErrorsCollection(ignoredErrors)) {
                reason = "ignored errors with unsupported metadata";
                return false;
            }

            foreach (IgnoredError ignoredError in ignoredErrors.Elements<IgnoredError>()) {
                if (!TryCreateIgnoredError(ignoredError, out _, out reason)) {
                    return false;
                }
            }

            return true;
        }

        internal static bool TryCreateHeaderPayload(ExcelSheet sheet, out byte[]? payload) {
            payload = null;
            if (!GetWorksheetIgnoredErrors(sheet).Any()) {
                return false;
            }

            payload = BuildHeaderPayload();
            return true;
        }

        internal static IReadOnlyList<byte[]> CreateIgnoredErrorPayloads(ExcelSheet sheet) {
            var payloads = new List<byte[]>();
            foreach (IgnoredError ignoredError in GetWorksheetIgnoredErrors(sheet)) {
                if (TryCreateIgnoredError(ignoredError, out IgnoredErrorFeature? feature, out _)) {
                    payloads.Add(BuildPayload(feature!));
                }
            }

            return payloads;
        }

        private static IReadOnlyList<IgnoredError> GetWorksheetIgnoredErrors(ExcelSheet sheet) {
            IgnoredErrors? ignoredErrors = GetWorksheetIgnoredErrorsCollection(sheet);
            return ignoredErrors == null
                ? Array.Empty<IgnoredError>()
                : ignoredErrors.Elements<IgnoredError>().ToArray();
        }

        private static IgnoredErrors? GetWorksheetIgnoredErrorsCollection(ExcelSheet sheet) {
            return sheet.WorksheetPart.Worksheet?.Elements<IgnoredErrors>().FirstOrDefault();
        }

        private static bool TryCreateIgnoredError(IgnoredError ignoredError, out IgnoredErrorFeature? feature, out string? reason) {
            feature = null;
            reason = null;

            if (ignoredError.HasChildren) {
                reason = "ignored errors with extension metadata";
                return false;
            }

            if (!SupportsKnownAttributes(ignoredError)) {
                reason = "ignored errors with unsupported metadata";
                return false;
            }

            if (ignoredError.CalculatedColumn?.Value == true) {
                reason = "ignored calculated-column errors";
                return false;
            }

            if (!TryParseRanges(ignoredError.SequenceOfReferences?.InnerText, out IReadOnlyList<CellRange> ranges, out reason)) {
                return false;
            }

            uint flags = BuildFlags(ignoredError);
            if (flags == 0) {
                reason = "ignored errors without supported error flags";
                return false;
            }

            feature = new IgnoredErrorFeature(ranges, flags);
            return true;
        }

        private static bool SupportsIgnoredErrorsCollection(IgnoredErrors ignoredErrors) {
            if (ignoredErrors.ChildElements.Any(child => child is not IgnoredError)) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in ignoredErrors.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                if (!string.Equals(attribute.LocalName, "count", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsKnownAttributes(OpenXmlElement element) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (!string.IsNullOrEmpty(attribute.NamespaceUri)) {
                    return false;
                }

                switch (attribute.LocalName) {
                    case "sqref":
                    case "evalError":
                    case "twoDigitTextYear":
                    case "numberStoredAsText":
                    case "formula":
                    case "formulaRange":
                    case "unlockedFormula":
                    case "emptyCellReference":
                    case "listDataValidation":
                    case "calculatedColumn":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool TryParseRanges(string? sequenceOfReferences, out IReadOnlyList<CellRange> ranges, out string? reason) {
            ranges = Array.Empty<CellRange>();
            reason = null;
            if (string.IsNullOrWhiteSpace(sequenceOfReferences)) {
                reason = "ignored errors without references";
                return false;
            }

            string[] parts = sequenceOfReferences!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0 || parts.Length > 1027) {
                reason = "ignored error reference counts outside BIFF8 limits";
                return false;
            }

            var parsed = new List<CellRange>(parts.Length);
            foreach (string part in parts) {
                string reference = part.Replace("$", string.Empty);
                if (!TryParseReference(reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    reason = "ignored error references";
                    return false;
                }

                if (firstRow < 1 || firstColumn < 1 || lastRow > 65536 || lastColumn > 256) {
                    reason = "ignored error references outside BIFF8 worksheet limits";
                    return false;
                }

                parsed.Add(new CellRange(
                    checked((ushort)(firstRow - 1)),
                    checked((ushort)(lastRow - 1)),
                    checked((ushort)(firstColumn - 1)),
                    checked((ushort)(lastColumn - 1))));
            }

            ranges = parsed;
            return true;
        }

        private static bool TryParseReference(string reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) {
            if (A1.TryParseRange(reference, out firstRow, out firstColumn, out lastRow, out lastColumn)) {
                return true;
            }

            (int row, int column) = A1.ParseCellRef(reference);
            if (row <= 0 || column <= 0) {
                firstRow = firstColumn = lastRow = lastColumn = 0;
                return false;
            }

            firstRow = lastRow = row;
            firstColumn = lastColumn = column;
            return true;
        }

        private static uint BuildFlags(IgnoredError ignoredError) {
            uint flags = 0;
            SetBit(ref flags, 0, ignoredError.EvalError?.Value == true);
            SetBit(ref flags, 1, ignoredError.EmptyCellReference?.Value == true);
            SetBit(ref flags, 2, ignoredError.NumberStoredAsText?.Value == true);
            SetBit(ref flags, 3, ignoredError.FormulaRange?.Value == true);
            SetBit(ref flags, 4, ignoredError.Formula?.Value == true);
            SetBit(ref flags, 5, ignoredError.TwoDigitTextYear?.Value == true);
            SetBit(ref flags, 6, ignoredError.UnlockedFormula?.Value == true);
            SetBit(ref flags, 7, ignoredError.ListDataValidation?.Value == true);
            return flags;
        }

        private static byte[] BuildHeaderPayload() {
            using var stream = new MemoryStream();
            WriteUInt16(stream, FeatHdrRecordType);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt16(stream, IsfIgnoredErrors);
            stream.WriteByte(1);
            WriteUInt32(stream, 0);
            return stream.ToArray();
        }

        private static byte[] BuildPayload(IgnoredErrorFeature feature) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, FeatRecordType);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt16(stream, IsfIgnoredErrors);
            stream.WriteByte(0);
            WriteUInt32(stream, 0);
            WriteUInt16(stream, checked((ushort)feature.Ranges.Count));
            WriteUInt32(stream, 4);
            WriteUInt16(stream, 0);

            foreach (CellRange range in feature.Ranges) {
                WriteCellRange(stream, range);
            }

            WriteUInt32(stream, feature.Flags);
            return stream.ToArray();
        }

        private static void WriteCellRange(Stream stream, CellRange range) {
            WriteUInt16(stream, range.FirstRow);
            WriteUInt16(stream, range.LastRow);
            WriteUInt16(stream, range.FirstColumn);
            WriteUInt16(stream, range.LastColumn);
        }

        private static void SetBit(ref uint flags, int bit, bool value) {
            if (value) {
                flags |= 1U << bit;
            }
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

        private sealed class IgnoredErrorFeature {
            internal IgnoredErrorFeature(IReadOnlyList<CellRange> ranges, uint flags) {
                Ranges = ranges;
                Flags = flags;
            }

            internal IReadOnlyList<CellRange> Ranges { get; }

            internal uint Flags { get; }
        }

        private readonly struct CellRange {
            internal CellRange(ushort firstRow, ushort lastRow, ushort firstColumn, ushort lastColumn) {
                FirstRow = firstRow;
                LastRow = lastRow;
                FirstColumn = firstColumn;
                LastColumn = lastColumn;
            }

            internal ushort FirstRow { get; }

            internal ushort LastRow { get; }

            internal ushort FirstColumn { get; }

            internal ushort LastColumn { get; }
        }
    }
}
