using System;
using System.IO;
using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO;

/// <summary>Bounded compound-directory inspection shared by legacy import adapters.</summary>
internal static class OfficeLegacyCompoundInspector {
    internal static OfficeLegacyInertContentKind Inspect(byte[] data, OfficeLegacyImportLimits limits, out bool inspectionIncomplete, CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        inspectionIncomplete = false;
        if (!OfficeLegacyImportBuffer.StartsWith(data, 0xD0, 0xCF, 0x11, 0xE0)) return OfficeLegacyInertContentKind.None;
        using var stream = new MemoryStream(data, writable: false);
        if (!OfficeCompoundFileReader.TryInspectDirectory(stream, limits.MaxInputBytes, limits.MaxCompoundStreams, cancellationToken, out var entries, out _)) {
            inspectionIncomplete = true;
            return OfficeLegacyInertContentKind.None;
        }

        OfficeLegacyInertContentKind result = OfficeLegacyInertContentKind.None;
        foreach (OfficeCompoundFileEntry entry in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            string path = entry.Path;
            if (Contains(path, "VBA") || Contains(path, "MACRO") || Contains(path, "SCRIPT")) {
                result |= OfficeLegacyInertContentKind.Macros | OfficeLegacyInertContentKind.EmbeddedCode;
            }
            if (Contains(path, "OBJECTPOOL") || Contains(path, "EMBEDDING") || Contains(path, "OLE10NATIVE")) {
                result |= OfficeLegacyInertContentKind.EmbeddedObjects;
            }
            if (Contains(path, "EXTERNALLINK") || Contains(path, "CONNECTION") || Contains(path, "LINKINFO") || Contains(path, "DDE")) {
                result |= OfficeLegacyInertContentKind.ExternalLinks;
            }
        }
        return result;
    }

    private static bool Contains(string value, string candidate) => value.IndexOf(candidate, StringComparison.OrdinalIgnoreCase) >= 0;
}
