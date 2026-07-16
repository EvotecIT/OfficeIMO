using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptPreservingWriter {
        private static bool TryRewriteClientDataStyle9(
            LegacyPptRecord clientData, byte[]? style9Record,
            out byte[] bytes) {
            if (clientData.Version != 0x0F || clientData.Instance != 0
                || clientData.Type != OfficeArtClientData) {
                bytes = clientData.CopyRecordBytes();
                return false;
            }
            var children = new List<byte[]>(clientData.Children.Count + 1);
            bool foundPpt9 = false;
            foreach (LegacyPptRecord child in clientData.Children) {
                if (child.Type != LegacyPptWriter.RecordProgTags) {
                    children.Add(child.CopyRecordBytes());
                    continue;
                }
                if (!TryRewriteShapeProgTags(child, style9Record,
                        ref foundPpt9, out byte[]? rewritten)) {
                    bytes = clientData.CopyRecordBytes();
                    return false;
                }
                if (rewritten != null) children.Add(rewritten);
            }
            if (!foundPpt9 && style9Record != null) {
                children.Add(LegacyPptWriter
                    .BuildShapePpt9ProgrammableTagsRecord(style9Record));
            }
            bytes = BuildRecord(clientData.Version, clientData.Instance,
                clientData.Type, Concat(children));
            return true;
        }

        private static bool TryRewriteShapeProgTags(
            LegacyPptRecord progTags, byte[]? style9Record,
            ref bool foundPpt9, out byte[]? bytes) {
            if (progTags.Version != 0x0F || progTags.Instance != 0) {
                bytes = progTags.CopyRecordBytes();
                return false;
            }
            var children = new List<byte[]>(progTags.Children.Count);
            foreach (LegacyPptRecord child in progTags.Children) {
                if (!IsShapePpt9BinaryTag(child)) {
                    children.Add(child.CopyRecordBytes());
                    continue;
                }
                if (foundPpt9) {
                    bytes = progTags.CopyRecordBytes();
                    return false;
                }
                foundPpt9 = true;
                if (style9Record != null) {
                    children.Add(LegacyPptWriter
                        .BuildShapePpt9BinaryTagRecord(style9Record));
                }
            }
            bytes = children.Count == 0
                ? null
                : BuildRecord(progTags.Version, progTags.Instance,
                    progTags.Type, Concat(children));
            return true;
        }

        private static bool IsShapePpt9BinaryTag(
            LegacyPptRecord record) {
            if (record.Version != 0x0F || record.Instance != 0
                || record.Type != LegacyPptWriter.RecordProgBinaryTag) {
                return false;
            }
            LegacyPptRecord[] names = record.Children.Where(child =>
                child.Type == RecordCString && child.Instance == 0)
                .ToArray();
            if (names.Length != 1 || names[0].Version != 0
                || (names[0].PayloadLength & 1) != 0) return false;
            try {
                return string.Equals(names[0].ReadUtf16Text()
                        .TrimEnd('\0'), "___PPT9",
                    StringComparison.Ordinal);
            } catch (InvalidDataException) {
                return false;
            }
        }
    }
}
