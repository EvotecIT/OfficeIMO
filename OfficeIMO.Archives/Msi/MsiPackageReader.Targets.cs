using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Archives;

internal sealed partial class MsiPackageReader {
    /// <summary>Plans deterministic logical targets without discarding any File/Media source member.</summary>
    private void PlanTargets(List<MsiPayloadFile> files) {
        var paths = new ArchivePaths(_limits);
        long retainedBytes = 0;
        foreach (var group in files.GroupBy(file => file.Path.Normalize(NormalizationForm.FormC), StringComparer.OrdinalIgnoreCase)) {
            _token.ThrowIfCancellationRequested();
            MsiPayloadFile[] aliases = group.ToArray();
            string canonical = aliases.Select(file => file.Path).OrderBy(path => path, StringComparer.Ordinal).First();
            long length = aliases[0].Length;
            foreach (MsiPayloadFile alias in aliases) {
                if (!IsEquivalentTarget(canonical, alias.Path))
                    throw new InvalidDataException("MSI target paths collide by directory case or Unicode: " + alias.Path);
                if (alias.Length != length) throw new InvalidDataException("MSI alias payload lengths conflict: " + canonical);
            }
            paths.Add(canonical);
            bool selected = _limits.Select(canonical);
            var target = new MsiPayloadTarget(canonical, aliases.Length, selected);
            foreach (MsiPayloadFile alias in aliases) {
                // Alias equality needs bytes even when the logical target is excluded.
                // Count every retained source row, never only the deduplicated output.
                if (selected || aliases.Length > 1)
                    ArchiveLimits.Add(ref retainedBytes, alias.Length, _limits.Expanded, "Selected and compared MSI content");
                alias.Target = target;
            }
        }
    }

    private static bool IsEquivalentTarget(string first, string second) {
        if (string.Equals(first, second, StringComparison.Ordinal)) return true;
        int separator = first.LastIndexOf('/');
        if (first.Length != second.Length || separator != second.LastIndexOf('/') ||
            !string.Equals(first.Substring(0, separator + 1), second.Substring(0, separator + 1), StringComparison.Ordinal)) return false;
        for (int index = separator + 1; index < first.Length; index++) {
            char left = first[index], right = second[index];
            if (left == right) continue;
            if (left >= 'A' && left <= 'Z') left = (char)(left + ('a' - 'A'));
            if (right >= 'A' && right <= 'Z') right = (char)(right + ('a' - 'A'));
            if (left < 'a' || left > 'z' || left != right) return false;
        }
        return true;
    }
}

/// <summary>A complete MSI alias group; actual byte equality is verified before publication.</summary>
internal sealed class MsiPayloadTarget {
    internal MsiPayloadTarget(string path, int aliasCount, bool selected) { Path = path; AliasCount = aliasCount; Selected = selected; }
    internal string Path { get; }
    internal int AliasCount { get; }
    internal bool Selected { get; }
}
