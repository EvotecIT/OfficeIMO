using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Validates complete VBA compound storages used by binary and Open XML presentations.</summary>
    internal static class LegacyPptVbaProjectCodec {
        internal static bool IsValidProject(byte[] bytes, out string? reason) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (!OfficeCompoundFileReader.TryRead(bytes,
                    out OfficeCompoundFile? compound, out reason)
                || compound == null) {
                return false;
            }
            if (!compound.Streams.ContainsKey("VBA/dir")
                || !compound.Streams.ContainsKey("VBA/_VBA_PROJECT")) {
                reason = "The compound storage has no VBA/dir or VBA/_VBA_PROJECT stream.";
                return false;
            }
            reason = null;
            return true;
        }
    }
}
