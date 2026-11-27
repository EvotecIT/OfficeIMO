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

        internal const int MaximumProjectBytes = 64 * 1024 * 1024;

        internal bool TryGetChange(LegacyPptPackageView view,
            out byte[]? currentBytes, out bool changed) {
            ArgumentNullException.ThrowIfNull(view);
            currentBytes = null;
            changed = false;
            LegacyPptPartView part = view.Presentation.RelatedPart(
                LegacyPptPackageView.VbaProjectRelationshipType);
            if (!part.IsNone) {
                try {
                    if (view.TryGetStoredLength(part.Name, out long length)
                        && length > MaximumProjectBytes) {
                        return false;
                    }
                    currentBytes = part.GetBytes();
                    if (currentBytes.Length > MaximumProjectBytes
                        || !LegacyPptVbaProjectCodec.IsValidProject(
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
