using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    // Layout, conversion and PDF output ask the same question for every text fragment.
    // These answers depend only on the registered faces, so they live until the next change.
    private const int MaximumCachedSelections = 4096;
    private const int MaximumCachedPlanLength = 1024;

    private SelectionCache? _selections;

    private List<OfficeFontFace> FacesToChange {
        get {
            Volatile.Write(ref _selections, null);
            return _faces;
        }
    }

    private List<string> FallbackFamiliesToChange {
        get {
            Volatile.Write(ref _selections, null);
            return _fallbackFamilies;
        }
    }

    private SelectionCache Selections {
        get {
            SelectionCache? current = Volatile.Read(ref _selections);
            if (current != null) return current;
            var created = new SelectionCache();
            return Interlocked.CompareExchange(ref _selections, created, null) ?? created;
        }
    }

    private string[] ParseFamilies(string familyNames) {
        ConcurrentDictionary<string, string[]> cache = Selections.ParsedFamilies;
        if (cache.TryGetValue(familyNames, out string[]? parsed)) return parsed;
        parsed = OfficeFontFamilyParser.Parse(familyNames).ToArray();
        if (cache.Count < MaximumCachedSelections) cache.TryAdd(familyNames, parsed);
        return parsed;
    }

    private FallbackSelection GetFallbackSelection(string familyNames, OfficeFontFaceDescriptor descriptor) {
        ConcurrentDictionary<(string, OfficeFontFaceDescriptor), FallbackSelection> cache = Selections.Fallbacks;
        (string, OfficeFontFaceDescriptor) key = (familyNames, descriptor);
        if (cache.TryGetValue(key, out FallbackSelection? selection)) return selection;
        selection = new FallbackSelection(ResolveFallbackCandidates(familyNames, descriptor), ParseFamilies(familyNames));
        if (cache.Count < MaximumCachedSelections) cache.TryAdd(key, selection);
        return selection;
    }

    private sealed class SelectionCache {
        internal readonly ConcurrentDictionary<string, string[]> ParsedFamilies =
            new ConcurrentDictionary<string, string[]>(StringComparer.Ordinal);
        internal readonly ConcurrentDictionary<(string, OfficeFontFaceDescriptor), (OfficeFontFace Face, int RegistrationIndex)[]> FamilyCandidates =
            new ConcurrentDictionary<(string, OfficeFontFaceDescriptor), (OfficeFontFace Face, int RegistrationIndex)[]>();
        internal readonly ConcurrentDictionary<(string, OfficeFontFaceDescriptor), FallbackSelection> Fallbacks =
            new ConcurrentDictionary<(string, OfficeFontFaceDescriptor), FallbackSelection>();
    }

    /// <summary>The ordered candidates for one family list and descriptor, with their resolved answers.</summary>
    private sealed class FallbackSelection {
        internal FallbackSelection(IReadOnlyList<OfficeFontFace> candidates, string[] requestedFamilies) {
            Candidates = candidates;
            ExplicitlySelectedFaces = new HashSet<string>(requestedFamilies, StringComparer.OrdinalIgnoreCase);
            ResolvedFamilyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < candidates.Count; index++) ResolvedFamilyNames.Add(candidates[index].ResourceFamilyName);
        }

        internal IReadOnlyList<OfficeFontFace> Candidates { get; }
        internal HashSet<string> ExplicitlySelectedFaces { get; }
        internal HashSet<string> ResolvedFamilyNames { get; }
        internal ConcurrentDictionary<string, (string Family, OfficeFontFace? Face)> Elements { get; } =
            new ConcurrentDictionary<string, (string Family, OfficeFontFace? Face)>(StringComparer.Ordinal);
        internal ConcurrentDictionary<string, IReadOnlyList<OfficeFontFallbackRun>> Plans { get; } =
            new ConcurrentDictionary<string, IReadOnlyList<OfficeFontFallbackRun>>(StringComparer.Ordinal);
        internal ConcurrentDictionary<string, OfficeFontFace?> FacesForText { get; } =
            new ConcurrentDictionary<string, OfficeFontFace?>(StringComparer.Ordinal);
    }
}
