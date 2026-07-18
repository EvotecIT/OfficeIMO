using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Retains the VBA persist mapping and projected part bytes for preservation-aware edits.</summary>
    internal sealed class LegacyPptVbaProjectProjection {
        private readonly byte[]? _projectBytes;

        private LegacyPptVbaProjectProjection(uint? persistId,
            byte[]? projectBytes) {
            PersistId = persistId;
            _projectBytes = projectBytes == null
                ? null
                : (byte[])projectBytes.Clone();
        }

        internal uint? PersistId { get; }

        internal static LegacyPptVbaProjectProjection Create(
            LegacyPptVbaProject? source) => new(
                source?.PersistId, source?.GetBytes());

        internal bool TryGetChange(PresentationPart? presentationPart,
            out byte[]? currentBytes, out bool changed) {
            currentBytes = null;
            changed = false;
            VbaProjectPart? part = presentationPart?.VbaProjectPart;
            if (part != null) {
                try {
                    using Stream stream = part.GetStream(FileMode.Open,
                        FileAccess.Read);
                    currentBytes = OfficeStreamReader.ReadAllBytes(stream,
                        64 * 1024 * 1024);
                    if (!LegacyPptVbaProjectCodec.IsValidProject(
                            currentBytes, out _)) {
                        return false;
                    }
                } catch (Exception exception) when (exception is IOException
                                                    or InvalidDataException
                                                    or UnauthorizedAccessException) {
                    return false;
                }
            }

            changed = !BytesEqual(_projectBytes, currentBytes);
            return true;
        }

        private static bool BytesEqual(byte[]? left, byte[]? right) {
            if (ReferenceEquals(left, right)) return true;
            if (left == null || right == null || left.Length != right.Length) {
                return false;
            }
            for (int index = 0; index < left.Length; index++) {
                if (left[index] != right[index]) return false;
            }
            return true;
        }
    }
}
