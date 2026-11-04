using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    /// <summary>Standard host locations, including the fonts shipped with installed Office apps on macOS.</summary>
    private static IEnumerable<string> CandidateFontDirectories(CancellationToken cancellationToken) {
        string windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (!string.IsNullOrEmpty(windows)) yield return Path.Combine(windows, "Fonts");
        yield return "/System/Library/Fonts";
        yield return "/System/Library/Fonts/Supplemental";
        yield return "/Library/Fonts";
        string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) {
            if (!string.IsNullOrEmpty(home)) {
                foreach (string directory in EnumerateUserFontDirectories(Path.Combine(home, "Library", "Fonts"), cancellationToken))
                    yield return directory;
            }
            foreach (string directory in OfficeInstalledApplicationFonts.GetDirectories()) yield return directory;
        }
        yield return "/usr/share/fonts/truetype/dejavu";
        yield return "/usr/share/fonts/truetype/liberation2";
        yield return "/usr/share/fonts/truetype/liberation";
        yield return "/usr/share/fonts/dejavu";
        yield return "/usr/share/fonts/liberation";
        yield return "/usr/share/fonts/truetype/crosextra";
        yield return "/usr/share/fonts/truetype/carlito";
        yield return "/usr/share/fonts/truetype/caladea";
        yield return "/usr/share/fonts/truetype/inter";
        if (!string.IsNullOrEmpty(home)) {
            foreach (string root in new[] { Path.Combine(home, ".fonts"), Path.Combine(home, ".local", "share", "fonts") })
                foreach (string directory in EnumerateUserFontDirectories(root, cancellationToken)) yield return directory;
        }
    }
    private static IEnumerable<string> SafeEnumerateFontFiles(string directory) {
        foreach (string file in OfficeCloudFontFiles.EnumerateChildren(directory, directories: false, maximum: 8192)) {
            string extension = Path.GetExtension(file);
            if (string.Equals(extension, ".ttf", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".otf", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".ttc", StringComparison.OrdinalIgnoreCase)) yield return file;
        }
    }

    /// <summary>Includes conventional nested user font folders without following links or scanning an unbounded tree.</summary>
    internal static IEnumerable<string> EnumerateUserFontDirectories(string root, CancellationToken cancellationToken = default) {
        const int MaximumDirectories = 512;
        const int MaximumDepth = 8;
        var pending = new Queue<(string Path, int Depth)>();
        pending.Enqueue((root, 0));
        int discovered = 1;
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var current = pending.Dequeue();
            yield return current.Path;
            if (current.Depth >= MaximumDepth || discovered >= MaximumDirectories) continue;
            foreach (string child in OfficeCloudFontFiles.EnumerateChildren(current.Path, directories: true,
                maximum: MaximumDirectories)) {
                cancellationToken.ThrowIfCancellationRequested();
                pending.Enqueue((child, current.Depth + 1));
                if (++discovered >= MaximumDirectories) break;
            }
        }
    }
}
