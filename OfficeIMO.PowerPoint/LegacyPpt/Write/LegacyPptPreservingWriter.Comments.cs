using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private const ushort RecordCString = 0x0FBA;
        private const ushort RecordComment10Atom = 0x2EE1;

        private static bool TryRewriteProgrammableTags(LegacyPptRecord progTags,
            IReadOnlyList<LegacyPptWriter.LegacyPptWriterComment> comments,
            out byte[]? bytes) {
            bytes = null;
            if (progTags.Version != 0x0F || progTags.Instance != 0) return false;
            var children = new List<byte[]>(progTags.Children.Count + 1);
            bool foundPpt10 = false;
            foreach (LegacyPptRecord child in progTags.Children) {
                if (!IsPpt10BinaryTag(child)) {
                    children.Add(child.CopyRecordBytes());
                    continue;
                }
                if (foundPpt10 || !TryRewritePpt10BinaryTag(child, comments,
                        out byte[]? rewrittenTag)) {
                    return false;
                }
                foundPpt10 = true;
                if (rewrittenTag != null) children.Add(rewrittenTag);
            }
            if (!foundPpt10 && comments.Count > 0) {
                children.Add(LegacyPptWriter.BuildPpt10BinaryTagRecord(comments));
            }
            if (children.Count > 0) {
                bytes = BuildRecord(progTags.Version, progTags.Instance, progTags.Type,
                    Concat(children));
            }
            return true;
        }

        private static bool TryRewritePpt10BinaryTag(LegacyPptRecord binaryTag,
            IReadOnlyList<LegacyPptWriter.LegacyPptWriterComment> comments,
            out byte[]? bytes) {
            bytes = null;
            if (binaryTag.Version != 0x0F || binaryTag.Instance != 0) return false;
            LegacyPptRecord[] dataBlobs = binaryTag.Children.Where(child =>
                child.Type == LegacyPptWriter.RecordBinaryTagDataBlob).ToArray();
            if (dataBlobs.Length != 1) return false;
            LegacyPptRecord dataBlob = dataBlobs[0];
            if (dataBlob.Version != 0 || dataBlob.Instance != 0) return false;
            IReadOnlyList<LegacyPptRecord> records;
            try {
                byte[] source = dataBlob.CopyRecordBytes();
                records = LegacyPptRecordReader.ReadSequence(source, 8,
                    dataBlob.PayloadLength, new LegacyPptImportOptions());
            } catch (InvalidDataException) {
                return false;
            }

            var dataRecords = new List<byte[]>(records.Count + comments.Count);
            bool insertedComments = false;
            foreach (LegacyPptRecord record in records) {
                if (record.Type == LegacyPptWriter.RecordComment10) {
                    if (!IsReplaceableCommentRecord(record)) return false;
                    if (!insertedComments) {
                        dataRecords.AddRange(LegacyPptWriter.BuildCommentRecords(comments));
                        insertedComments = true;
                    }
                    continue;
                }
                dataRecords.Add(record.CopyRecordBytes());
            }
            if (!insertedComments) {
                dataRecords.AddRange(LegacyPptWriter.BuildCommentRecords(comments));
            }
            if (dataRecords.Count == 0) return true;

            byte[] rewrittenBlob = BuildRecord(dataBlob.Version, dataBlob.Instance,
                dataBlob.Type, Concat(dataRecords));
            var children = new List<byte[]>(binaryTag.Children.Count);
            foreach (LegacyPptRecord child in binaryTag.Children) {
                children.Add(ReferenceEquals(child, dataBlob)
                    ? rewrittenBlob
                    : child.CopyRecordBytes());
            }
            bytes = BuildRecord(binaryTag.Version, binaryTag.Instance, binaryTag.Type,
                Concat(children));
            return true;
        }

        private static bool IsReplaceableCommentRecord(LegacyPptRecord record) {
            if (record.Version != 0x0F || record.Instance != 0) return false;
            int atomCount = 0;
            var stringInstances = new HashSet<ushort>();
            foreach (LegacyPptRecord child in record.Children) {
                if (child.Type == RecordComment10Atom) {
                    if (++atomCount != 1 || child.Version != 0 || child.Instance != 0
                        || child.PayloadLength != 28 || !IsValidCommentAtom(child)) {
                        return false;
                    }
                } else if (child.Type == RecordCString) {
                    int maximumLength = child.Instance == 1 ? 64000 : 104;
                    if (child.Version != 0 || child.Instance > 2
                        || child.PayloadLength > maximumLength
                        || (child.PayloadLength & 1) != 0
                        || !stringInstances.Add(child.Instance)) {
                        return false;
                    }
                } else {
                    return false;
                }
            }
            return atomCount == 1;
        }

        private static bool IsValidCommentAtom(LegacyPptRecord atom) {
            if (atom.ReadInt32(0) < 0) return false;
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
            return year >= 1 && year <= 9999 && month >= 1 && month <= 12
                && day >= 1 && day <= DateTime.DaysInMonth(year, month)
                && dayOfWeek <= 6 && hour <= 23 && minute <= 59 && second <= 59
                && milliseconds <= 999;
        }

        private static bool IsPpt10BinaryTag(LegacyPptRecord record) {
            if (record.Type != LegacyPptWriter.RecordProgBinaryTag) return false;
            LegacyPptRecord? name = record.Children.FirstOrDefault(child =>
                child.Type == RecordCString && child.Instance == 0);
            if (name == null || (name.PayloadLength & 1) != 0) return false;
            try {
                return string.Equals(name.ReadUtf16Text().TrimEnd('\0'), "___PPT10",
                    StringComparison.Ordinal);
            } catch (InvalidDataException) {
                return false;
            }
        }
    }
}
