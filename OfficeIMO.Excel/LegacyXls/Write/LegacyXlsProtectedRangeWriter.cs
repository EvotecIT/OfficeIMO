using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsProtectedRangeWriter {
        private const ushort FeatRecordType = 0x0868;
        private const ushort IsfProtection = 0x0002;
        private const int BiffMaxRecordDataLength = 8224;

        internal static bool SupportsWorksheetProtectedRanges(ExcelSheet sheet, out string? reason) {
            reason = null;
            ProtectedRanges? protectedRanges = sheet.WorksheetPart.Worksheet?.Elements<ProtectedRanges>().FirstOrDefault();
            if (protectedRanges == null) {
                return true;
            }

            if (!SupportsProtectedRangesCollection(protectedRanges)) {
                reason = "protected ranges with unsupported metadata";
                return false;
            }

            if (!protectedRanges.Elements<ProtectedRange>().Any()) {
                return true;
            }

            foreach (ProtectedRange protectedRange in protectedRanges.Elements<ProtectedRange>()) {
                if (!TryCreateProtectedRange(protectedRange, out _, out reason)) {
                    return false;
                }
            }

            return true;
        }

        internal static IReadOnlyList<byte[]> CreateProtectedRangePayloads(ExcelSheet sheet) {
            ProtectedRanges? protectedRanges = sheet.WorksheetPart.Worksheet?.Elements<ProtectedRanges>().FirstOrDefault();
            if (protectedRanges == null) {
                return Array.Empty<byte[]>();
            }

            var payloads = new List<byte[]>();
            foreach (ProtectedRange protectedRange in protectedRanges.Elements<ProtectedRange>()) {
                if (TryCreateProtectedRange(protectedRange, out ProtectedRangeFeature? feature, out _)) {
                    payloads.Add(BuildPayload(feature!));
                }
            }

            return payloads;
        }

        private static bool TryCreateProtectedRange(ProtectedRange protectedRange, out ProtectedRangeFeature? feature, out string? reason) {
            feature = null;
            reason = null;

            string? name = protectedRange.Name?.Value;
            if (string.IsNullOrWhiteSpace(name)) {
                reason = "protected ranges without names";
                return false;
            }

            if (name!.Length > ushort.MaxValue) {
                reason = "protected ranges with names longer than BIFF8 limits";
                return false;
            }

            if (HasAnyAttribute(protectedRange, "algorithmName", "hashValue", "saltValue", "spinCount")) {
                reason = "protected ranges with modern password hashes";
                return false;
            }

            if (HasAnyAttribute(protectedRange, "securityDescriptor")) {
                reason = "protected ranges with security descriptors";
                return false;
            }

            if (protectedRange.HasChildren || !SupportsKnownAttributes(protectedRange)) {
                reason = "protected ranges with unsupported metadata";
                return false;
            }

            if (!TryParseRanges(protectedRange.SequenceOfReferences?.InnerText, out IReadOnlyList<CellRange> ranges, out reason)) {
                return false;
            }

            if (!SupportsProtectedRangePayloadLength(name, ranges.Count, out reason)) {
                return false;
            }

            ushort? passwordHash = null;
            string? rawPassword = protectedRange.Password?.Value;
            if (!string.IsNullOrWhiteSpace(rawPassword)) {
                if (!ushort.TryParse(rawPassword!.Trim(), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out ushort hash)) {
                    reason = "protected ranges with non-legacy password hashes";
                    return false;
                }

                passwordHash = hash;
            }

            feature = new ProtectedRangeFeature(name, ranges, passwordHash);
            return true;
        }

        private static bool SupportsProtectedRangePayloadLength(string name, int rangeCount, out string? reason) {
            reason = null;
            long payloadLength = 35L
                + (8L * rangeCount)
                + GetUnicodeStringPayloadLength(name);
            if (payloadLength > BiffMaxRecordDataLength) {
                reason = "protected range payload lengths outside BIFF8 limits";
                return false;
            }

            return true;
        }

        private static long GetUnicodeStringPayloadLength(string text) {
            return 3L + (CanUseCompressedString(text) ? text.Length : 2L * text.Length);
        }

        private static bool TryParseRanges(string? sequenceOfReferences, out IReadOnlyList<CellRange> ranges, out string? reason) {
            ranges = Array.Empty<CellRange>();
            reason = null;
            if (string.IsNullOrWhiteSpace(sequenceOfReferences)) {
                reason = "protected ranges without references";
                return false;
            }

            string[] parts = sequenceOfReferences!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0 || parts.Length > 432) {
                reason = "protected range counts outside BIFF8 limits";
                return false;
            }

            var parsed = new List<CellRange>(parts.Length);
            foreach (string part in parts) {
                string rangeText = part.Replace("$", string.Empty);
                if (!A1.TryParseRange(rangeText, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    reason = "protected range references";
                    return false;
                }

                if (firstRow < 1 || firstColumn < 1 || lastRow > 65536 || lastColumn > 256) {
                    reason = "protected range references outside BIFF8 worksheet limits";
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

        private static byte[] BuildPayload(ProtectedRangeFeature feature) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, FeatRecordType);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt16(stream, IsfProtection);
            stream.WriteByte(0);
            WriteUInt32(stream, 0);
            WriteUInt16(stream, checked((ushort)feature.Ranges.Count));
            WriteUInt32(stream, 0);
            WriteUInt16(stream, 0);

            foreach (CellRange range in feature.Ranges) {
                WriteCellRange(stream, range);
            }

            WriteUInt32(stream, 0);
            WriteUInt32(stream, feature.PasswordHash ?? 0);
            WriteUnicodeString(stream, feature.Name);
            return stream.ToArray();
        }

        private static void WriteCellRange(Stream stream, CellRange range) {
            WriteUInt16(stream, range.FirstRow);
            WriteUInt16(stream, range.LastRow);
            WriteUInt16(stream, range.FirstColumn);
            WriteUInt16(stream, range.LastColumn);
        }

        private static void WriteUnicodeString(Stream stream, string text) {
            byte[] encoded = EncodeUnicodeString(text, out byte flags);
            WriteUInt16(stream, checked((ushort)text.Length));
            stream.WriteByte(flags);
            stream.Write(encoded, 0, encoded.Length);
        }

        private static byte[] EncodeUnicodeString(string text, out byte flags) {
            if (CanUseCompressedString(text)) {
                flags = 0;
                return Encoding.ASCII.GetBytes(text);
            }

            flags = 1;
            return Encoding.Unicode.GetBytes(text);
        }

        private static bool CanUseCompressedString(string text) {
            for (int i = 0; i < text.Length; i++) {
                if (text[i] > 0x7f) {
                    return false;
                }
            }

            return true;
        }

        private static bool SupportsProtectedRangesCollection(ProtectedRanges protectedRanges) {
            if (protectedRanges.ChildElements.Any(child => child is not ProtectedRange)) {
                return false;
            }

            foreach (OpenXmlAttribute attribute in protectedRanges.GetAttributes()) {
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
                    case "name":
                    case "sqref":
                    case "password":
                    case "algorithmName":
                    case "hashValue":
                    case "saltValue":
                    case "spinCount":
                    case "securityDescriptor":
                        break;
                    default:
                        return false;
                }
            }

            return true;
        }

        private static bool HasAnyAttribute(OpenXmlElement element, params string[] localNames) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                for (int i = 0; i < localNames.Length; i++) {
                    if (string.Equals(attribute.LocalName, localNames[i], StringComparison.Ordinal)) {
                        return true;
                    }
                }
            }

            return false;
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

        private sealed class ProtectedRangeFeature {
            internal ProtectedRangeFeature(string name, IReadOnlyList<CellRange> ranges, ushort? passwordHash) {
                Name = name;
                Ranges = ranges;
                PasswordHash = passwordHash;
            }

            internal string Name { get; }

            internal IReadOnlyList<CellRange> Ranges { get; }

            internal ushort? PasswordHash { get; }
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
