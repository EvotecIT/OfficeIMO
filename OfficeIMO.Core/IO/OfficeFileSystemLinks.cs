using System;
using System.IO;

namespace OfficeIMO.Internal;

/// <summary>Tells links and reparse points apart from ordinary entries without reading the entry's attributes on Unix.</summary>
/// <remarks>
/// Reading <see cref="FileSystemInfo.Attributes"/>, <see cref="File.GetAttributes(string)"/> or
/// <see cref="FileInfo.Length"/> also computes the read-only flag. For an entry the process doesn't own, such as a
/// system font, .NET then checks the process's group membership, and on iOS that check throws
/// <see cref="OverflowException"/>. Off Windows a link is what makes an entry a reparse point, so the link target alone
/// answers the question. Windows keeps the attribute check, which also covers junctions and other reparse points.
/// </remarks>
internal static class OfficeFileSystemLinks {
    /// <summary>Whether the entry is a symbolic link, junction or other reparse point.</summary>
    internal static bool IsLinkOrReparsePoint(FileSystemInfo entry) {
        if (entry == null) throw new ArgumentNullException(nameof(entry));
#if NET6_0_OR_GREATER
        if (!OperatingSystem.IsWindows()) {
            // With a trailing separator, LinkTarget looks at the folder a link points to, not at the link.
            string path = entry.FullName;
            string trimmed = Path.TrimEndingDirectorySeparator(path);
            if (trimmed.Length != path.Length) entry = entry is DirectoryInfo ? new DirectoryInfo(trimmed) : new FileInfo(trimmed);
            return entry.LinkTarget != null;
        }
#endif
        return (entry.Attributes & FileAttributes.ReparsePoint) != 0;
    }

    /// <summary>Whether the path names a symbolic link, junction or other reparse point.</summary>
    internal static bool IsLinkOrReparsePoint(string path) => IsLinkOrReparsePoint(new FileInfo(path));
}
