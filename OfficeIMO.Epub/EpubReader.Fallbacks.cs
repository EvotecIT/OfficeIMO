using System.Threading;

namespace OfficeIMO.Epub;

internal static partial class EpubReader {
    // Cache chain resolution separately from parsed payloads so repeated positions and
    // shared fallback tails do not repeatedly traverse the manifest graph.
    private static ManifestItem? ResolveChapterResource(EpubPackage package, ManifestItem primary,
        Dictionary<string, ManifestItem?> selections, EpubDiagnosticCollector diagnostics, CancellationToken token) {
        if (IsChapterManifestItem(primary)) return primary;
        if (selections.TryGetValue(primary.Id, out ManifestItem? cached)) return cached;
        ManifestItem current = primary;
        var chain = new List<ManifestItem>();
        var visited = new HashSet<string>(StringComparer.Ordinal);
        ManifestItem? selected;
        while (true) {
            token.ThrowIfCancellationRequested();
            if (IsChapterManifestItem(current)) { selected = current; break; }
            if (selections.TryGetValue(current.Id, out selected)) break;
            if (!visited.Add(current.Id)) {
                diagnostics.Warning("epub.spine.fallback-cycle", "Skipped spine resource with a cyclic fallback chain.", primary.FullPath);
                selected = null; break;
            }
            chain.Add(current);
            if (current.FallbackId == null) {
                diagnostics.Warning("epub.spine.unsupported-media-type",
                    $"Skipped spine resource '{primary.FullPath}' with unsupported media type '{primary.MediaType}'.",
                    primary.FullPath, primary.MediaType);
                selected = null; break;
            }
            if (!package.Manifest.TryGetValue(current.FallbackId, out ManifestItem? next)) {
                diagnostics.Warning("epub.spine.fallback-missing",
                    $"Skipped spine resource '{primary.FullPath}' because fallback id '{current.FallbackId}' is missing.", primary.FullPath);
                selected = null; break;
            }
            current = next;
        }
        foreach (ManifestItem item in chain) selections[item.Id] = selected;
        return selected;
    }
}
