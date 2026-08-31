using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO;

/// <summary>Bounded compound-directory inspection shared by legacy import adapters.</summary>
internal static class OfficeLegacyCompoundInspector {
    private static readonly byte[] CompoundSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    internal static bool IsValidCompound(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        return TryInspectDirectory(data, limits, cancellationToken, out _);
    }

    internal static OfficeLegacyInertContentKind Inspect(byte[] data, OfficeLegacyImportLimits limits, out bool inspectionIncomplete, CancellationToken cancellationToken = default) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        inspectionIncomplete = HasCompoundSignature(data);
        if (!TryInspectDirectory(data, limits, cancellationToken, out var entries)) {
            return OfficeLegacyInertContentKind.None;
        }
        inspectionIncomplete = false;

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

    private static bool TryInspectDirectory(byte[] data, OfficeLegacyImportLimits limits, CancellationToken cancellationToken, out IReadOnlyList<OfficeCompoundFileEntry> entries) {
        entries = Array.Empty<OfficeCompoundFileEntry>();
        if (!HasCompoundSignature(data) || data.LongLength > limits.MaxInputBytes) return false;
        using var stream = new MemoryStream(data, writable: false);
        return OfficeCompoundFileReader.TryInspectDirectory(stream, limits.MaxInputBytes, limits.MaxCompoundStreams, cancellationToken, out entries, out _);
    }

    private static bool HasCompoundSignature(byte[] data) =>
        data.Length >= CompoundSignature.Length && OfficeLegacyImportBuffer.StartsWith(data, CompoundSignature);

    private static bool Contains(string value, string candidate) => value.IndexOf(candidate, StringComparison.OrdinalIgnoreCase) >= 0;
}
