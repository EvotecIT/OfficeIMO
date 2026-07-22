using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Diagnostics;

namespace OfficeIMO.PowerPoint.LegacyPpt {
    public sealed partial class LegacyPptPresentation {
        private const ushort RecordProgTags = 0x1388;
        private const ushort RecordProgBinaryTag = 0x138A;
        private const ushort RecordBinaryTagDataBlob = 0x138B;
        private const ushort RecordComment10 = 0x2EE0;
        private const ushort RecordComment10Atom = 0x2EE1;
        private const string Ppt10TagName = "___PPT10";

        private void ParseComments(LegacyPptRecord slideRecord, LegacyPptSlide slide,
            LegacyPptImportOptions options) {
            foreach (LegacyPptRecord progTags in slideRecord.Children.Where(record =>
                         record.Type == RecordProgTags)) {
                foreach (LegacyPptRecord binaryTag in progTags.Children.Where(record =>
                             record.Type == RecordProgBinaryTag)) {
                    LegacyPptRecord? tagName = binaryTag.Children.FirstOrDefault(record =>
                        record.Type == RecordCString && record.Instance == 0);
                    if (tagName == null || !TryReadUnicodeString(tagName, out string name)
                        || !string.Equals(name, Ppt10TagName, StringComparison.Ordinal)) {
                        continue;
                    }

                    LegacyPptRecord[] dataBlobs = binaryTag.Children.Where(record =>
                        record.Type == RecordBinaryTagDataBlob).ToArray();
                    if (dataBlobs.Length != 1) {
                        AddDiagnostic("PPT-COMMENT-TAG-DATA", LegacyPptDiagnosticSeverity.Warning,
                            $"Slide {slide.SlideId} has a malformed PP10 programmable tag; comments remain preserve-only.",
                            binaryTag.Offset);
                        continue;
                    }
                    ParseCommentDataBlob(dataBlobs[0], slide, options);
                }
            }
        }

        private void ParseCommentDataBlob(LegacyPptRecord dataBlob, LegacyPptSlide slide,
            LegacyPptImportOptions options) {
            IReadOnlyList<LegacyPptRecord> records;
            try {
                byte[] bytes = dataBlob.CopyRecordBytes();
                records = LegacyPptRecordReader.ReadSequence(bytes, 8,
                    dataBlob.PayloadLength, options, _recordBudget);
            } catch (InvalidDataException) {
                AddDiagnostic("PPT-COMMENT-DATA-TRUNCATED", LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slide.SlideId} has malformed PP10 tag data; comments remain preserve-only.",
                    dataBlob.Offset);
                return;
            }

            foreach (LegacyPptRecord record in records.Where(record =>
                         record.Type == RecordComment10)) {
                if (_commentCount >= options.MaxCommentCount) {
                    throw new InvalidDataException(
                        $"The binary PowerPoint comment count exceeds {options.MaxCommentCount}.");
                }
                _commentCount++;
                if (TryReadComment(record, slide.SlideId, out LegacyPptComment? comment)
                    && comment != null) {
                    slide.AddComment(comment);
                }
            }
        }

        private bool TryReadComment(LegacyPptRecord container, uint slideId,
            out LegacyPptComment? comment) {
            comment = null;
            if (container.Version != 0x0F || container.Instance != 0) {
                AddDiagnostic("PPT-COMMENT-CONTAINER", LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slideId} has a malformed Comment10Container; the comment remains preserve-only.",
                    container.Offset);
                return false;
            }
            LegacyPptRecord? atom = container.Children.FirstOrDefault(record =>
                record.Type == RecordComment10Atom);
            if (atom == null || atom.Version != 0 || atom.Instance != 0
                || atom.PayloadLength != 28
                || container.Children.Count(record => record.Type == RecordComment10Atom) != 1) {
                AddDiagnostic("PPT-COMMENT-ATOM", LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slideId} has a malformed Comment10Atom; the comment remains preserve-only.",
                    container.Offset);
                return false;
            }

            int index = atom.ReadInt32(0);
            if (index < 0) {
                AddDiagnostic("PPT-COMMENT-INDEX", LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slideId} has a negative comment index; the comment remains preserve-only.",
                    atom.Offset);
                return false;
            }

            if (!TryReadCommentDate(atom, out DateTime? createdAtUtc)) {
                AddDiagnostic("PPT-COMMENT-DATETIME", LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slideId} has an invalid comment creation time; the comment remains preserve-only.",
                    atom.Offset);
                return false;
            }

            if (!TryReadCommentString(container, instance: 0, maximumBytes: 104,
                    out string author)
                || !TryReadCommentString(container, instance: 1, maximumBytes: 64000,
                    out string text)
                || !TryReadCommentString(container, instance: 2, maximumBytes: 104,
                    out string initials)) {
                AddDiagnostic("PPT-COMMENT-STRING", LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slideId} has a malformed comment author, text, or initials string; the comment remains preserve-only.",
                    container.Offset);
                return false;
            }
            if (container.Children.Any(child => child.Type != RecordComment10Atom
                    && (child.Type != RecordCString || child.Instance > 2))) {
                AddDiagnostic("PPT-COMMENT-EXTENSION", LegacyPptDiagnosticSeverity.Warning,
                    $"Slide {slideId} has comment extension data that is retained only in the binary source.",
                    container.Offset);
            }
            comment = new LegacyPptComment(index, author, initials, text, createdAtUtc,
                atom.ReadInt32(20), atom.ReadInt32(24));
            return true;
        }

        private static bool TryReadCommentString(LegacyPptRecord container, ushort instance,
            int maximumBytes, out string value) {
            value = string.Empty;
            LegacyPptRecord[] records = container.Children.Where(child =>
                child.Type == RecordCString && child.Instance == instance).ToArray();
            if (records.Length == 0) return true;
            if (records.Length != 1 || records[0].Version != 0
                || records[0].PayloadLength > maximumBytes) {
                return false;
            }
            return TryReadUnicodeString(records[0], out value);
        }

        private static bool TryReadUnicodeString(LegacyPptRecord record, out string value) {
            value = string.Empty;
            if ((record.PayloadLength & 1) != 0) return false;
            try {
                value = record.ReadUtf16Text().TrimEnd('\0');
                return true;
            } catch (InvalidDataException) {
                return false;
            }
        }

        private static bool TryReadCommentDate(LegacyPptRecord atom, out DateTime? value) {
            value = null;
            ushort year = atom.ReadUInt16(4);
            ushort month = atom.ReadUInt16(6);
            ushort dayOfWeek = atom.ReadUInt16(8);
            ushort day = atom.ReadUInt16(10);
            ushort hour = atom.ReadUInt16(12);
            ushort minute = atom.ReadUInt16(14);
            ushort second = atom.ReadUInt16(16);
            ushort milliseconds = atom.ReadUInt16(18);
            if (year == 0 && month == 0 && dayOfWeek == 0 && day == 0 && hour == 0
                && minute == 0 && second == 0 && milliseconds == 0) {
                return true;
            }
            if (year < 1 || year > 9999 || month < 1 || month > 12 || day < 1
                || day > DateTime.DaysInMonth(year, month) || hour > 23 || minute > 59
                || second > 59 || milliseconds > 999 || dayOfWeek > 6) {
                return false;
            }
            value = new DateTime(year, month, day, hour, minute, second, milliseconds,
                DateTimeKind.Utc);
            return true;
        }
    }
}
