using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordDocInfoList = 0x07D0;
        private const ushort RecordVbaInfo = 0x03FF;
        private const ushort RecordVbaInfoAtom = 0x0400;
        private const ushort RecordExternalOleObjectStorage = 0x1011;

        internal static bool TryReadVbaProject(
            PowerPointPresentation presentation, out byte[]? projectBytes,
            out string? reason) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            projectBytes = null;
            reason = null;
            VbaProjectPart? part = presentation.OpenXmlDocument
                .PresentationPart?.VbaProjectPart;
            if (part == null) return true;
            if (part.Parts.Any()) {
                reason = "Related VBA signature or data parts do not have a PowerPoint 97-2003 representation.";
                return false;
            }
            try {
                using Stream stream = part.GetStream(FileMode.Open,
                    FileAccess.Read);
                projectBytes = OfficeStreamReader.ReadAllBytes(stream,
                    64 * 1024 * 1024);
            } catch (Exception exception) when (exception is IOException
                                                or InvalidDataException
                                                or UnauthorizedAccessException) {
                reason = $"The VBA project part could not be read: {exception.Message}";
                return false;
            }
            if (!LegacyPptVbaProjectCodec.IsValidProject(projectBytes,
                    out reason)) {
                projectBytes = null;
                return false;
            }
            return true;
        }

        internal static byte[] BuildVbaProjectStorageRecord(
            byte[] projectBytes) {
            if (projectBytes == null) throw new ArgumentNullException(nameof(projectBytes));
            if (!LegacyPptVbaProjectCodec.IsValidProject(projectBytes,
                    out string? reason)) {
                throw new InvalidDataException(reason
                    ?? "The VBA project is not a valid compound storage.");
            }
            return BuildRecord(version: 0, instance: 0,
                RecordExternalOleObjectStorage, projectBytes);
        }

        internal static byte[] RewriteDocumentVbaInfo(
            LegacyPptRecord document, uint? persistId) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var children = new List<byte[]>(document.Children.Count + 1);
            bool foundDocInfo = false;
            foreach (LegacyPptRecord child in document.Children) {
                if (child.Type != RecordDocInfoList) {
                    children.Add(child.CopyRecordBytes());
                    continue;
                }
                foundDocInfo = true;
                var docInfoChildren = child.Children
                    .Where(item => item.Type != RecordVbaInfo)
                    .Select(item => item.CopyRecordBytes()).ToList();
                if (persistId.HasValue) {
                    docInfoChildren.Add(BuildVbaInfoContainer(
                        persistId.Value));
                }
                children.Add(BuildRecord(child.Version, child.Instance,
                    child.Type, Concat(docInfoChildren)));
            }
            if (!foundDocInfo && persistId.HasValue) {
                children.Add(BuildContainer(RecordDocInfoList, instance: 0,
                    new[] { BuildVbaInfoContainer(persistId.Value) }));
            }
            return BuildRecord(document.Version, document.Instance,
                document.Type, Concat(children));
        }

        private static byte[] BuildVbaInfoContainer(uint persistId) {
            var payload = new byte[12];
            WriteUInt32(payload, 0, persistId);
            WriteUInt32(payload, 4, 1);
            WriteUInt32(payload, 8, 2);
            byte[] atom = BuildRecord(version: 2, instance: 0,
                RecordVbaInfoAtom, payload);
            return BuildContainer(RecordVbaInfo, instance: 1,
                new[] { atom });
        }
    }
}
