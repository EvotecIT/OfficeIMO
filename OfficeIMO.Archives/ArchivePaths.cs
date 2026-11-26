using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives;

internal sealed class ArchivePaths {
    private readonly HashSet<string> _files = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _directories = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _explicitDirectories = new(StringComparer.OrdinalIgnoreCase);
    private readonly ArchiveLimits _limits;
    private readonly long? _maximumIndexBytes;
    private long _indexBytes;

    internal ArchivePaths(ArchiveLimits limits, long? maximumIndexBytes = null) {
        _limits = limits;
        _maximumIndexBytes = maximumIndexBytes;
    }
    internal string Add(string authored) {
        string path = Validate(authored);
        string key = path.Normalize(NormalizationForm.FormC);
        if (_files.Contains(key) || _directories.ContainsKey(key)) throw new InvalidDataException("Archive member paths collide: " + authored);
        ChargeIndex(key, 0);
        _files.Add(key);
        AddParents(path, authored);
        return path;
    }
    /// <summary>Registers an explicit directory record, such as a ZIP entry ending in '/', without publishing a file.</summary>
    internal string AddDirectory(string authored) {
        string path = Validate(authored);
        string key = path.Normalize(NormalizationForm.FormC);
        if (_explicitDirectories.Contains(key)) throw new InvalidDataException("Archive directory records collide: " + authored);
        ChargeIndex(key, 0);
        _explicitDirectories.Add(key);
        AddParents(path + "/", authored);
        return path;
    }
    private string Validate(string authored) {
        string path = authored.Replace('\\', '/');
        if (OfficeArchiveSafety.IsUnsafePath(path))
            throw new InvalidDataException("Archive member has an unsafe or oversized relative path: " + authored);
        string[] parts = path.Split('/');
        foreach (string part in parts) ValidateSegment(part);
        ArchiveLimits.Check(path.Length, _limits.PathLength, "MaxPathLength");
        ArchiveLimits.Check(parts.Length, _limits.Depth, "MaxPathDepth");
        return path;
    }
    private void AddParents(string path, string authored) {
        int separator = path.IndexOf('/');
        while (separator >= 0) {
            string parent = path.Substring(0, separator);
            string parentKey = parent.Normalize(NormalizationForm.FormC);
            if (_files.Contains(parentKey)) throw new InvalidDataException("An archive file is also a parent directory: " + authored);
            if (_directories.TryGetValue(parentKey, out string? existing)) {
                if (!string.Equals(existing, parent, StringComparison.Ordinal)) throw new InvalidDataException("Archive directory paths collide by case or Unicode normalization: " + authored);
            } else {
                ChargeIndex(parentKey, parent.Length);
                _directories.Add(parentKey, parent);
            }
            separator = path.IndexOf('/', separator + 1);
        }
    }
    private void ChargeIndex(string key, int originalLength) {
        if (!_maximumIndexBytes.HasValue) return;
        // Count retained UTF-16 strings and reserve space per hash-table entry, including resizes.
        long bytes = 128 + 2L * (key.Length + (long)originalLength);
        ArchiveLimits.Add(ref _indexBytes, bytes, _maximumIndexBytes.Value, nameof(ArchiveExtractionOptions.MaxMetadataBytes));
    }

    internal static void ValidateSegment(string part) {
        if (part.Length == 0 || part.Length > 255 || Encoding.UTF8.GetByteCount(part) > 255)
            throw new InvalidDataException("Archive member has an empty or oversized path segment.");
        int dot = part.IndexOf('.');
        // Win32 ignores spaces before the extension when matching devices, so "CON .txt" is CON.
        string baseName = (dot < 0 ? part : part.Substring(0, dot)).TrimEnd(' ');
        if (part.EndsWith(".", StringComparison.Ordinal)
            || part.EndsWith(" ", StringComparison.Ordinal) || part.Any(c => c < 32 || ":<>\"|?*/\\".IndexOf(c) >= 0)
            || IsDevice(baseName)) throw new InvalidDataException("Archive member has a nonportable path segment: '" + part + "'");
    }
    /// <summary>Rejects a linked node or links inside a controlled subtree, trusting caller and OS ancestors.</summary>
    internal static void RejectLinks(string fullPath, string? controlledRoot = null) {
        string target = NormalizeLinkPath(fullPath);
        string boundary = NormalizeLinkPath(controlledRoot ?? fullPath);
        StringComparison comparison = Path.DirectorySeparatorChar == '\\' ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        string prefix = boundary.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            ? boundary : boundary + Path.DirectorySeparatorChar;
        if (!string.Equals(target, boundary, comparison) && !target.StartsWith(prefix, comparison))
            throw new IOException("Extraction paths must remain inside the controlled staging directory.");
        string? current = target;
        while (!string.IsNullOrEmpty(current)) {
            try {
                if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("Extraction paths must not contain symbolic links or reparse points: " + current);
            } catch (FileNotFoundException) { } catch (DirectoryNotFoundException) { }
            if (string.Equals(current, boundary, comparison)) break;
            current = Path.GetDirectoryName(current);
        }
    }
    private static string NormalizeLinkPath(string path) {
        string full = Path.GetFullPath(path);
        int rootLength = Path.GetPathRoot(full)!.Length;
        string trimmed = full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return trimmed.Length < rootLength ? full.Substring(0, rootLength) : trimmed;
    }
    private static bool IsDevice(string value) => value.Equals("CON", StringComparison.OrdinalIgnoreCase)
        || value.Equals("CONIN$", StringComparison.OrdinalIgnoreCase) || value.Equals("CONOUT$", StringComparison.OrdinalIgnoreCase)
        || value.Equals("PRN", StringComparison.OrdinalIgnoreCase) || value.Equals("AUX", StringComparison.OrdinalIgnoreCase)
        || value.Equals("NUL", StringComparison.OrdinalIgnoreCase)
        || value.Length == 4 && (value.StartsWith("COM", StringComparison.OrdinalIgnoreCase) || value.StartsWith("LPT", StringComparison.OrdinalIgnoreCase))
            // Windows also reserves the ISO-8859-1 superscript digits 1, 2 and 3 in these device names.
            && (value[3] >= '1' && value[3] <= '9' || value[3] == '\u00b9' || value[3] == '\u00b2' || value[3] == '\u00b3');
}
