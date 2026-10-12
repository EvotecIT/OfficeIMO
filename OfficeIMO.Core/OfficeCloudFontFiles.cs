using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>Finds locally cached Office fonts without downloading or installing them.</summary>
internal static class OfficeCloudFontFiles {
    /// <summary>Finds existing versioned cache roots beneath the current user's local application data.</summary>
    internal static IEnumerable<string> GetRoots() {
        string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(local)) yield break;
        foreach (string root in EnumerateRoots(Path.Combine(local, "Microsoft", "FontCache"))) yield return root;
    }

    /// <summary>Enumerates ordinary cloud-font folders within a bounded local cache.</summary>
    internal static IEnumerable<string> EnumerateRoots(string cacheDirectory) {
        foreach (string version in EnumerateChildren(cacheDirectory, directories: true, maximum: 32)) {
            string root = Path.Combine(version, "CloudFonts");
            if (IsOrdinaryDirectory(root)) yield return root;
        }
    }

    /// <summary>Finds candidate files for a normalized family; loaders retain authority over their metadata.</summary>
    internal static IEnumerable<string> GetFamilyFiles(string normalizedFamilyKey) {
        foreach (string root in GetRoots()) {
            foreach (string file in EnumerateFamilyFiles(root, normalizedFamilyKey)) yield return file;
        }
    }

    internal static IEnumerable<string> EnumerateFamilyFiles(string cloudRoot, string normalizedFamilyKey) {
        foreach (string family in EnumerateChildren(cloudRoot, directories: true, maximum: 512)) {
            if (!string.Equals(FamilyKey(Path.GetFileName(family)), normalizedFamilyKey, StringComparison.Ordinal)) continue;
            // Office cache filenames can be numeric. The folder supplies a
            // candidate hint; the font loader must still validate its family.
            foreach (string file in EnumerateChildren(family, directories: false, maximum: 128)) {
                string extension = Path.GetExtension(file);
                if (string.Equals(extension, ".ttf", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(extension, ".otf", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(extension, ".ttc", StringComparison.OrdinalIgnoreCase)) yield return file;
            }
        }
    }

    private static string FamilyKey(string value) {
        var result = new StringBuilder(value.Length);
        foreach (char c in value) if (char.IsLetterOrDigit(c)) result.Append(char.ToLowerInvariant(c));
        return result.ToString();
    }

    private static bool IsOrdinaryDirectory(string path) {
        try {
            var info = new DirectoryInfo(path);
            return info.Exists && (info.Attributes & FileAttributes.ReparsePoint) == 0;
        } catch (Exception exception) when (IsPathFailure(exception)) {
            return false;
        }
    }

    private static IEnumerable<string> EnumerateChildren(string root, bool directories, int maximum) {
        if (!IsOrdinaryDirectory(root)) yield break;
        IEnumerator<string>? entries = null;
        try {
            entries = (directories ? Directory.EnumerateDirectories(root) : Directory.EnumerateFiles(root)).GetEnumerator();
        } catch (Exception exception) when (IsPathFailure(exception)) {
            // An unavailable local cache does not prevent font fallback.
        }
        if (entries == null) yield break;
        using (entries) {
            for (int inspected = 0; inspected < maximum; inspected++) {
                string path;
                try {
                    if (!entries.MoveNext()) break;
                    path = entries.Current;
                    if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0) continue;
                } catch (Exception exception) when (IsPathFailure(exception)) {
                    break;
                }
                yield return path;
            }
        }
    }

    private static bool IsPathFailure(Exception exception) => exception is IOException ||
        exception is UnauthorizedAccessException || exception is ArgumentException || exception is NotSupportedException;
}
