using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordSlideNumberMetaCharacter = 0x0FD8;
        private const ushort RecordDateTimeMetaCharacter = 0x0FF7;
        private const ushort RecordGenericDateMetaCharacter = 0x0FF8;
        private const ushort RecordHeaderMetaCharacter = 0x0FF9;
        private const ushort RecordFooterMetaCharacter = 0x0FFA;
        private const ushort RecordRtfDateTimeMetaCharacter = 0x1015;

        private IReadOnlyList<LegacyPptTextField> ReadTextFields(
            LegacyPptRecord? textbox, string text,
            LegacyPptImportOptions options, out bool hasFieldRecords,
            out bool isMalformed) {
            LegacyPptRecord[] records = textbox?.Children.Where(record =>
                IsTextFieldRecord(record.Type)).ToArray()
                ?? Array.Empty<LegacyPptRecord>();
            hasFieldRecords = records.Length > 0;
            isMalformed = false;
            var result = new List<LegacyPptTextField>(records.Length);
            var positions = new HashSet<int>();
            foreach (LegacyPptRecord record in records) {
                if (!TryReadTextField(record, text, out LegacyPptTextField?
                        field) || field == null
                    || !positions.Add(field.Position)) {
                    isMalformed = true;
                    if (options.ReportUnsupportedContent) {
                        AddDiagnostic("PPT-TEXT-FIELD-MALFORMED",
                            LegacyPptDiagnosticSeverity.Warning,
                            "A dynamic text-field metacharacter is malformed, overlaps another field, or points outside its text body.",
                            record.Offset);
                    }
                    continue;
                }
                result.Add(field);
            }
            return result.OrderBy(field => field.Position).ToArray();
        }

        private static bool TryReadTextField(LegacyPptRecord record,
            string text, out LegacyPptTextField? field) {
            field = null;
            int expectedLength = record.Type switch {
                RecordDateTimeMetaCharacter => 8,
                RecordRtfDateTimeMetaCharacter => 132,
                _ => 4
            };
            if (record.Version != 0 || record.Instance != 0
                || record.PayloadLength != expectedLength) return false;
            uint rawPosition = record.ReadUInt32(0);
            if (rawPosition > int.MaxValue) return false;
            int position = unchecked((int)rawPosition);
            if (position < 0 || position >= text.Length
                || text[position] == '\n') return false;
            switch (record.Type) {
                case RecordSlideNumberMetaCharacter:
                    field = new LegacyPptTextField(position,
                        LegacyPptTextFieldKind.SlideNumber);
                    return true;
                case RecordDateTimeMetaCharacter:
                    byte format = record.ReadByte(4);
                    if (format > 0x0C) return false;
                    field = new LegacyPptTextField(position,
                        LegacyPptTextFieldKind.DateTime, format);
                    return true;
                case RecordGenericDateMetaCharacter:
                    field = new LegacyPptTextField(position,
                        LegacyPptTextFieldKind.GenericDate);
                    return true;
                case RecordHeaderMetaCharacter:
                    field = new LegacyPptTextField(position,
                        LegacyPptTextFieldKind.Header);
                    return true;
                case RecordFooterMetaCharacter:
                    field = new LegacyPptTextField(position,
                        LegacyPptTextFieldKind.Footer);
                    return true;
                case RecordRtfDateTimeMetaCharacter:
                    field = new LegacyPptTextField(position,
                        LegacyPptTextFieldKind.RtfDateTime,
                        rtfFormat: record.ReadUtf16Text(4, 128)
                            .TrimEnd('\0'));
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsTextFieldRecord(ushort type) => type
            is RecordSlideNumberMetaCharacter
                or RecordDateTimeMetaCharacter
                or RecordGenericDateMetaCharacter
                or RecordHeaderMetaCharacter
                or RecordFooterMetaCharacter
                or RecordRtfDateTimeMetaCharacter;
    }
}
