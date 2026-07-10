using System;
using System.IO.Compression;

namespace OfficeIMO.Shared.Packaging;

/// <summary>
/// Centralizes archive-entry safety rules shared by ZIP-backed OfficeIMO format owners.
/// </summary>
internal static class OfficeArchiveSafety {
    private static readonly char[] PathSeparators = { '/' };

    /// <summary>Normalizes an archive entry name to forward-slash notation.</summary>
    internal static string NormalizeEntryName(string? fullName) {
        string value = fullName ?? string.Empty;
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        string normalized = value.Replace('\\', '/').Trim();
        while (normalized.StartsWith("./", StringComparison.Ordinal)) {
            normalized = normalized.Substring(2);
        }

        return normalized;
    }

    /// <summary>Returns true when an entry name is absolute, traverses parents, or contains a NUL.</summary>
    internal static bool IsUnsafePath(string fullName) {
        if (fullName.Length == 0) return true;
        if (fullName[0] == '/' || fullName[0] == '\\') return true;
        if (ContainsNull(fullName)) return true;
        if (fullName.Length >= 2 && char.IsLetter(fullName[0]) && fullName[1] == ':') return true;

        string[] segments = fullName.Split(PathSeparators, StringSplitOptions.RemoveEmptyEntries);
        foreach (string segment in segments) {
            if (segment == "." || segment == "..") return true;
        }

        return false;
    }

    /// <summary>Returns the logical path depth of an archive entry.</summary>
    internal static int ComputeDepth(string fullName, bool isDirectory) {
        string normalized = isDirectory ? fullName.TrimEnd('/') : fullName;
        if (normalized.Length == 0) return 0;

        int depth = 1;
        for (int i = 0; i < normalized.Length; i++) {
            if (normalized[i] == '/') depth++;
        }

        return depth;
    }

    /// <summary>Reads an entry's declared uncompressed length without allowing metadata failures to escape.</summary>
    internal static bool TryGetLength(ZipArchiveEntry entry, out long length) {
        try {
            length = entry.Length;
            return true;
        } catch {
            length = 0;
            return false;
        }
    }

    /// <summary>Checks an entry's declared expansion ratio.</summary>
    internal static bool IsCompressionRatioExceeded(ZipArchiveEntry entry, long uncompressedLength, double maxRatio) {
        if (maxRatio <= 0 || uncompressedLength <= 0) return false;

        long compressedLength;
        try {
            compressedLength = entry.CompressedLength;
        } catch {
            return false;
        }

        if (compressedLength <= 0) return false;
        return (double)uncompressedLength / compressedLength > maxRatio;
    }

    private static bool ContainsNull(string value) {
        for (int i = 0; i < value.Length; i++) {
            if (value[i] == '\0') return true;
        }
        return false;
    }
}
