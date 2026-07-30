using OfficeIMO.Drawing;
using System.Collections.ObjectModel;

namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    private Dictionary<string, PdfEmbeddedFontFallbackCandidate[]>? _renderingProfileFamilyFallbacks;
    private PdfEmbeddedFontFallbackCandidate[]? _renderingProfileDeclaredFallbackCandidates;
    private HashSet<string>? _renderingProfileOwnedNamedFamilyNames;

    /// <summary>
    /// Applies the text and font resources from a shared OfficeIMO rendering profile.
    /// </summary>
    /// <remarks>
    /// PDF-specific pagination, conformance, metadata, and document security remain configured on
    /// <see cref="PdfOptions"/>. The profile supplies only format-neutral shaping, language, and fonts.
    /// </remarks>
    /// <param name="profile">Shared rendering profile.</param>
    /// <param name="mode">Whether profile-owned resources replace or overlay existing PDF settings.</param>
    /// <returns>This options instance for fluent configuration.</returns>
    public PdfOptions UseRenderingProfile(
        OfficeRenderingProfile profile,
        OfficeRenderingProfileApplyMode mode = OfficeRenderingProfileApplyMode.Replace) {
        Guard.NotNull(profile, nameof(profile));
        if (mode != OfficeRenderingProfileApplyMode.Replace
            && mode != OfficeRenderingProfileApplyMode.Overlay) {
            throw new ArgumentOutOfRangeException(nameof(mode));
        }

        OfficeFontFaceCollection profileFonts = profile.Fonts;
        PdfEmbeddedFontFallbackSet? existingFallbacks = mode == OfficeRenderingProfileApplyMode.Overlay
            ? EmbeddedFontFallbacksSnapshot?.Clone()
            : null;
        var profileOwnedFallbackNames = new HashSet<string>(
            _renderingProfileDeclaredFallbackCandidates?
                .Select(candidate => candidate.FontName)
                ?? Enumerable.Empty<string>(),
            StringComparer.OrdinalIgnoreCase);
        var preservedNamedFamilyNames = new HashSet<string>(
            _namedFontFamilies?.Values
                .Select(family => family.FamilyName)
                .Where(name =>
                    _renderingProfileOwnedNamedFamilyNames?.Contains(name) != true)
                ?? Enumerable.Empty<string>(),
            StringComparer.OrdinalIgnoreCase);
        ReadOnlyCollection<PdfEmbeddedFontFamily> families = CreateProfileFontFamilies(profileFonts);
        existingFallbacks = PromoteCompatibilitySlotFallbacks(
            existingFallbacks,
            families,
            preservedNamedFamilyNames,
            mode);
        var preservedFallbackNames = new HashSet<string>(
            existingFallbacks?.Candidates
                .Select(candidate => candidate.FontName)
                .Where(name => !profileOwnedFallbackNames.Contains(name))
                ?? Enumerable.Empty<string>(),
            StringComparer.OrdinalIgnoreCase);
        ValidateRenderingProfileFamilyCapacity(families, existingFallbacks, mode);
        ValidateOptionalLanguage(profile.TextShapingLanguage, nameof(profile));

        if (mode == OfficeRenderingProfileApplyMode.Replace || profile.TextShapingProvider != null) {
            TextShapingProvider = profile.TextShapingProvider;
        }
        if (mode == OfficeRenderingProfileApplyMode.Replace || profile.TextShapingLanguage != null) {
            Language = profile.TextShapingLanguage;
        }

        if (mode == OfficeRenderingProfileApplyMode.Replace) {
            ClearNamedFontFamilies();
            ClearEmbeddedStandardFontMappings();
            _embeddedFontFallbacks = null;
            _renderingProfileFamilyFallbacks?.Clear();
            _renderingProfileDeclaredFallbackCandidates = null;
            _renderingProfileOwnedNamedFamilyNames?.Clear();
        }
        if (families.Count > 0) {
            foreach (PdfEmbeddedFontFamily family in families) {
                if (mode == OfficeRenderingProfileApplyMode.Overlay
                    && preservedNamedFamilyNames.Contains(family.FamilyName)) {
                    continue;
                }
                RegisterRenderingProfileNamedFamily(
                    mode == OfficeRenderingProfileApplyMode.Overlay
                        ? MergeRenderingProfileNamedFamily(family, profileFonts)
                        : family);
            }
        }

        RegisterProfileFamilyFallbacks(
            profileFonts,
            mode == OfficeRenderingProfileApplyMode.Overlay
                ? preservedNamedFamilyNames
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase));

        PdfEmbeddedFontFallbackCandidate[] profileCandidates =
            CreateProfileFallbackCandidates(profileFonts);
        PdfEmbeddedFontFallbackCandidate[] inheritedFamilyRefreshCandidates =
            mode == OfficeRenderingProfileApplyMode.Overlay
                ? CreateProfileFallbackCandidates(
                    profileFonts,
                    EnumerateDeclaredFallbackFamilyNames(
                        _renderingProfileDeclaredFallbackCandidates))
                : Array.Empty<PdfEmbeddedFontFallbackCandidate>();
        PdfEmbeddedFontFallbackCandidate[] overlayCandidates =
            ConcatDistinctCandidateVariants(
                inheritedFamilyRefreshCandidates,
                profileCandidates);
        PdfEmbeddedFontFallbackCandidate[] profileCandidateVariants = mode == OfficeRenderingProfileApplyMode.Overlay
            ? overlayCandidates
                .Where(candidate =>
                    !preservedFallbackNames.Contains(candidate.FontName)
                    && !preservedNamedFamilyNames.Contains(candidate.FontName))
                .ToArray()
            : overlayCandidates;
        _renderingProfileDeclaredFallbackCandidates =
            _renderingProfileDeclaredFallbackCandidates == null
                ? profileCandidateVariants
                : OverlayDeclaredFallbackCandidateVariants(
                    _renderingProfileDeclaredFallbackCandidates,
                    profileCandidateVariants);
        PdfEmbeddedFontFallbackCandidate[] regularProfileCandidates =
            SelectRenderingProfileCandidates(
                _renderingProfileDeclaredFallbackCandidates,
                bold: false,
                italic: false);
        PdfEmbeddedFontFallbackCandidate[] combinedCandidates =
            MergeProfileFallbackCandidates(
                existingFallbacks,
                regularProfileCandidates,
                profileOwnedFallbackNames);
        if (combinedCandidates.Length > 0) {
            EnsureNamedFallbackCandidatesRegistered(combinedCandidates);

            // The named families above may include complete styled profile families.
            // Store the planner directly so registering its regular candidates cannot replace them.
            _embeddedFontFallbacks = new PdfEmbeddedFontFallbackSet(combinedCandidates);
        }

        return this;
    }

    internal bool TryGetRenderingProfileFamilyFallbacks(
        string? familyName,
        out PdfEmbeddedFontFallbackSet? fallbackSet) {
        fallbackSet = null;
        if (!TryGetRenderingProfileFamilyCandidates(
                familyName,
                out PdfEmbeddedFontFallbackCandidate[]? registered)) {
            return false;
        }

        PdfEmbeddedFontFallbackCandidate[] candidates =
            SelectRenderingProfileFamilyCandidates(
                registered!,
                bold: false,
                italic: false);
        if (candidates.Length == 0) {
            return false;
        }

        fallbackSet = new PdfEmbeddedFontFallbackSet(candidates);
        return true;
    }

    internal bool ShouldPreferSelectedCallerFamily(string? familyName) {
        if (!TryGetNamedFontFamily(
                familyName,
                out PdfEmbeddedFontFamily? selectedFamily)
            || selectedFamily == null) {
            return false;
        }
        return _renderingProfileOwnedNamedFamilyNames?.Contains(
            selectedFamily.FamilyName) != true;
    }

    internal bool TryGetEffectiveRenderingProfileFallbacks(
        string? familyName,
        bool bold,
        bool italic,
        out PdfEmbeddedFontFallbackSet? fallbackSet) {
        fallbackSet = null;
        if (!TryGetRenderingProfileFamilyCandidates(
                familyName,
                out PdfEmbeddedFontFallbackCandidate[]? registered)
            || registered == null) {
            return false;
        }

        PdfEmbeddedFontFallbackCandidate[] styledCandidates =
            SelectRenderingProfileFamilyCandidates(registered, bold, italic);
        PdfEmbeddedFontFallbackCandidate[] declaredCandidates =
            SelectRenderingProfileCandidates(
                _renderingProfileDeclaredFallbackCandidates
                    ?? Array.Empty<PdfEmbeddedFontFallbackCandidate>(),
                bold,
                italic);
        styledCandidates = MergeFallbackCandidates(
            styledCandidates,
            declaredCandidates);
        PdfEmbeddedFontFallbackCandidate[] combinedCandidates =
            MergeFallbackCandidates(styledCandidates, _embeddedFontFallbacks?.Candidates);
        if (combinedCandidates.Length == 0) {
            return false;
        }

        EnsureNamedFallbackCandidatesRegistered(combinedCandidates);
        fallbackSet = new PdfEmbeddedFontFallbackSet(combinedCandidates);
        return true;
    }

    private bool TryGetRenderingProfileFamilyCandidates(
        string? familyName,
        out PdfEmbeddedFontFallbackCandidate[]? candidates) {
        candidates = null;
        if (string.IsNullOrWhiteSpace(familyName)
            || _renderingProfileFamilyFallbacks == null) {
            return false;
        }

        var merged = new List<PdfEmbeddedFontFallbackCandidate>();
        var variants = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string familyCandidate in EnumerateOfficeFontFamilyCandidates(familyName!)) {
            if (_renderingProfileFamilyFallbacks.TryGetValue(
                    familyCandidate,
                    out PdfEmbeddedFontFallbackCandidate[]? registered)) {
                foreach (PdfEmbeddedFontFallbackCandidate candidate in registered) {
                    if (variants.Add(CandidateVariantKey(candidate))) {
                        merged.Add(candidate);
                    }
                }
            }
        }
        candidates = merged.ToArray();
        return candidates.Length > 0;
    }

    internal PdfEmbeddedFontFallbackSet? GetEffectiveRenderingProfileDeclaredFallbacks(
        bool bold,
        bool italic) {
        if (_renderingProfileDeclaredFallbackCandidates == null
            || _renderingProfileDeclaredFallbackCandidates.Length == 0) {
            return null;
        }

        PdfEmbeddedFontFallbackCandidate[] declaredCandidates =
            SelectRenderingProfileCandidates(
                _renderingProfileDeclaredFallbackCandidates,
                bold,
                italic);
        PdfEmbeddedFontFallbackCandidate[] combinedCandidates =
            MergeFallbackCandidates(
                declaredCandidates,
                _embeddedFontFallbacks?.Candidates);
        if (combinedCandidates.Length == 0) {
            return null;
        }

        EnsureNamedFallbackCandidatesRegistered(combinedCandidates);
        return new PdfEmbeddedFontFallbackSet(combinedCandidates);
    }

    private void RegisterProfileFamilyFallbacks(
        OfficeFontFaceCollection fonts,
        HashSet<string> excludedResourceFamilyNames) {
        var familiesToRefresh = new HashSet<string>(
            fonts.Faces
                .Select(face => face.FamilyName),
            StringComparer.OrdinalIgnoreCase);
        foreach (IGrouping<string, OfficeFontFace> family in fonts.Faces
            .Where(face => familiesToRefresh.Contains(face.FamilyName))
            .GroupBy(face => face.FamilyName, StringComparer.OrdinalIgnoreCase)) {
            PdfEmbeddedFontFallbackCandidate[] candidates = family
                .Where(face => !excludedResourceFamilyNames.Contains(face.ResourceFamilyName))
                .Where(face => !face.UnicodeRanges.IsAll)
                .Reverse()
                .Concat(family
                    .Where(face => !excludedResourceFamilyNames.Contains(face.ResourceFamilyName))
                    .Where(face => face.UnicodeRanges.IsAll)
                    .Reverse())
                .GroupBy(
                    face => face.ResourceFamilyName
                        + "\u001f"
                        + ((int)face.Style).ToString(
                            System.Globalization.CultureInfo.InvariantCulture),
                    StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .Select(face => new PdfEmbeddedFontFallbackCandidate(
                    face.ResourceFamilyName,
                    face.Data,
                    face.UnicodeRanges,
                    face.Style,
                    face.FamilyName))
                .ToArray();
            if (candidates.Length == 0) {
                continue;
            }

            Dictionary<string, PdfEmbeddedFontFallbackCandidate[]> fallbacks =
                _renderingProfileFamilyFallbacks ??=
                    new Dictionary<string, PdfEmbeddedFontFallbackCandidate[]>(
                        StringComparer.OrdinalIgnoreCase);
            PdfEmbeddedFontFallbackCandidate[] merged = fallbacks.TryGetValue(
                    family.Key,
                    out PdfEmbeddedFontFallbackCandidate[]? existing)
                ? OverlayFallbackCandidateVariants(existing, candidates)
                : candidates;
            fallbacks[family.Key] = merged;
        }
    }

    private static PdfEmbeddedFontFallbackCandidate[] SelectRenderingProfileFamilyCandidates(
        IReadOnlyList<PdfEmbeddedFontFallbackCandidate> candidates,
        bool bold,
        bool italic) {
        return SelectRenderingProfileCandidates(candidates, bold, italic);
    }

    private static PdfEmbeddedFontFallbackCandidate[] SelectRenderingProfileCandidates(
        IReadOnlyList<PdfEmbeddedFontFallbackCandidate> candidates,
        bool bold,
        bool italic) {
        OfficeFontStyle requested =
            (bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        var selected = new List<PdfEmbeddedFontFallbackCandidate>();
        foreach (IGrouping<string, PdfEmbeddedFontFallbackCandidate> family in candidates
            .GroupBy(candidate => candidate.PlannerFamilyName, StringComparer.OrdinalIgnoreCase)) {
            OfficeFontStyle selectedStyle = family.Any(item => item.Style == requested)
                ? requested
                : family.Any(item => item.Style == OfficeFontStyle.Regular)
                    ? OfficeFontStyle.Regular
                    : family.First().Style;
            selected.AddRange(family.Where(item => item.Style == selectedStyle));
        }
        return selected.ToArray();
    }

    private void EnsureNamedFallbackCandidatesRegistered(
        IEnumerable<PdfEmbeddedFontFallbackCandidate> candidates) {
        foreach (PdfEmbeddedFontFallbackCandidate candidate in candidates) {
            if (!HasNamedFontFamily(candidate.FontName)) {
                RegisterRenderingProfileNamedFamily(new PdfEmbeddedFontFamily(
                    candidate.FontName,
                    candidate.DataSnapshot));
            }
        }
    }

    private void RegisterRenderingProfileNamedFamily(PdfEmbeddedFontFamily family) {
        if (!TryRegisterNamedFontFamily(family)) {
            throw new InvalidOperationException(
                $"No more than {MaximumNamedFontFamilies} named font families can be registered.");
        }
        (_renderingProfileOwnedNamedFamilyNames ??=
            new HashSet<string>(StringComparer.OrdinalIgnoreCase))
            .Add(family.FamilyName);
    }

    private PdfEmbeddedFontFamily MergeRenderingProfileNamedFamily(
        PdfEmbeddedFontFamily supplied,
        OfficeFontFaceCollection profileFonts) {
        if (_renderingProfileOwnedNamedFamilyNames?.Contains(supplied.FamilyName) != true
            || !TryGetNamedFontFamilyDirect(
                supplied.FamilyName,
                out PdfEmbeddedFontFamily? existing)
            || existing == null) {
            return supplied;
        }

        OfficeFontFace[] suppliedFaces = profileFonts.Faces
            .Where(face => string.Equals(
                face.ResourceFamilyName,
                supplied.FamilyName,
                StringComparison.OrdinalIgnoreCase))
            .ToArray();
        OfficeFontFace? regular = SelectProfileFace(
            suppliedFaces,
            OfficeFontStyle.Regular);
        OfficeFontFace? bold = SelectProfileFace(
            suppliedFaces,
            OfficeFontStyle.Bold);
        OfficeFontFace? italic = SelectProfileFace(
            suppliedFaces,
            OfficeFontStyle.Italic);
        OfficeFontFace? boldItalic = SelectProfileFace(
            suppliedFaces,
            OfficeFontStyle.Bold | OfficeFontStyle.Italic);
        return new PdfEmbeddedFontFamily(
            supplied.FamilyName,
            regular?.Data ?? existing.RegularSnapshot,
            bold?.Data ?? existing.BoldSnapshot,
            italic?.Data ?? existing.ItalicSnapshot,
            boldItalic?.Data ?? existing.BoldItalicSnapshot);
    }

    private void ReleaseRenderingProfileFontOwnership(string familyName) {
        bool releasedProfileOwnedFamily =
            _renderingProfileOwnedNamedFamilyNames?.Remove(familyName) == true;
        if (_renderingProfileDeclaredFallbackCandidates != null) {
            _renderingProfileDeclaredFallbackCandidates =
                _renderingProfileDeclaredFallbackCandidates
                    .Where(candidate =>
                        !string.Equals(
                            candidate.FontName,
                            familyName,
                            StringComparison.OrdinalIgnoreCase))
                    .ToArray();
        }
        if (releasedProfileOwnedFamily && _embeddedFontFallbacks != null) {
            PdfEmbeddedFontFallbackCandidate[] candidates =
                _embeddedFontFallbacks.Candidates
                    .Where(candidate =>
                        !string.Equals(
                            candidate.FontName,
                            familyName,
                            StringComparison.OrdinalIgnoreCase))
                    .ToArray();
            if (candidates.Length != _embeddedFontFallbacks.Candidates.Count) {
                if (candidates.Length == 0) {
                    _embeddedFontFallbacks = null;
                } else if (_embeddedFontFallbacks.UsesNamedFontFamilies) {
                    _embeddedFontFallbacks =
                        new PdfEmbeddedFontFallbackSet(candidates);
                } else {
                    PdfStandardFont[] slots = _embeddedFontFallbacks.Candidates
                        .Select((candidate, index) => new {
                            Candidate = candidate,
                            Slot = _embeddedFontFallbacks.FontSlots[index]
                        })
                        .Where(entry =>
                            !string.Equals(
                                entry.Candidate.FontName,
                                familyName,
                                StringComparison.OrdinalIgnoreCase))
                        .Select(entry => entry.Slot)
                        .ToArray();
                    _embeddedFontFallbacks =
                        new PdfEmbeddedFontFallbackSet(candidates, slots);
                }
            }
        }
        if (_renderingProfileFamilyFallbacks == null) {
            return;
        }
        foreach (string key in _renderingProfileFamilyFallbacks.Keys.ToArray()) {
            PdfEmbeddedFontFallbackCandidate[] remaining =
                _renderingProfileFamilyFallbacks[key]
                    .Where(candidate =>
                        !string.Equals(
                            candidate.FontName,
                            familyName,
                            StringComparison.OrdinalIgnoreCase))
                    .ToArray();
            if (remaining.Length == 0) {
                _renderingProfileFamilyFallbacks.Remove(key);
            } else {
                _renderingProfileFamilyFallbacks[key] = remaining;
            }
        }
    }

    private void ValidateRenderingProfileFamilyCapacity(
        IEnumerable<PdfEmbeddedFontFamily> profileFamilies,
        PdfEmbeddedFontFallbackSet? promotedFallbacks,
        OfficeRenderingProfileApplyMode mode) {
        var familyKeys = new HashSet<string>(StringComparer.Ordinal);
        if (mode == OfficeRenderingProfileApplyMode.Overlay
            && _namedFontFamilies != null) {
            familyKeys.UnionWith(_namedFontFamilies.Keys);
        }
        foreach (PdfEmbeddedFontFamily family in profileFamilies) {
            familyKeys.Add(NormalizeNamedFontFamilyKey(family.FamilyName));
        }
        if (promotedFallbacks != null) {
            foreach (PdfEmbeddedFontFallbackCandidate candidate in promotedFallbacks.Candidates) {
                familyKeys.Add(NormalizeNamedFontFamilyKey(candidate.FontName));
            }
        }
        if (familyKeys.Count > MaximumNamedFontFamilies) {
            throw new InvalidOperationException(
                $"No more than {MaximumNamedFontFamilies} named font families can be registered.");
        }
    }

    private PdfEmbeddedFontFallbackSet? PromoteCompatibilitySlotFallbacks(
        PdfEmbeddedFontFallbackSet? fallbackSet,
        IEnumerable<PdfEmbeddedFontFamily> profileFamilies,
        HashSet<string> preservedNamedFamilyNames,
        OfficeRenderingProfileApplyMode mode) {
        if (fallbackSet == null || fallbackSet.UsesNamedFontFamilies) {
            return fallbackSet;
        }

        var prospectiveFamilies = new Dictionary<string, PdfEmbeddedFontFamily>(
            StringComparer.Ordinal);
        if (mode == OfficeRenderingProfileApplyMode.Overlay && _namedFontFamilies != null) {
            foreach (KeyValuePair<string, PdfEmbeddedFontFamily> entry in _namedFontFamilies) {
                prospectiveFamilies[entry.Key] = entry.Value;
            }
        }
        foreach (PdfEmbeddedFontFamily family in profileFamilies) {
            if (mode == OfficeRenderingProfileApplyMode.Overlay
                && preservedNamedFamilyNames.Contains(family.FamilyName)) {
                continue;
            }
            prospectiveFamilies[NormalizeNamedFontFamilyKey(family.FamilyName)] = family;
        }

        var promoted = new List<PdfEmbeddedFontFallbackCandidate>(
            fallbackSet.Candidates.Count);
        for (int index = 0; index < fallbackSet.Candidates.Count; index++) {
            PdfEmbeddedFontFallbackCandidate candidate = fallbackSet.Candidates[index];
            string familyName = candidate.FontName;
            string key = NormalizeNamedFontFamilyKey(familyName);
            if (prospectiveFamilies.TryGetValue(key, out PdfEmbeddedFontFamily? existing)
                && !NamedFamilyMatchesFallbackCandidate(existing, candidate)) {
                familyName = FindReusablePromotedFallbackFamilyName(
                        candidate,
                        fallbackSet.FontSlots[index],
                        prospectiveFamilies)
                    ?? CreatePromotedFallbackFamilyName(
                        candidate.FontName,
                        fallbackSet.FontSlots[index],
                        prospectiveFamilies);
                key = NormalizeNamedFontFamilyKey(familyName);
            }

            var promotedCandidate = new PdfEmbeddedFontFallbackCandidate(
                familyName,
                candidate.DataSnapshot,
                candidate.UnicodeRanges,
                candidate.Style,
                candidate.PlannerFamilyName);
            promoted.Add(promotedCandidate);
            if (!prospectiveFamilies.ContainsKey(key)) {
                prospectiveFamilies[key] = new PdfEmbeddedFontFamily(
                    familyName,
                    candidate.DataSnapshot);
            }
        }

        return new PdfEmbeddedFontFallbackSet(promoted);
    }

    private static bool NamedFamilyMatchesFallbackCandidate(
        PdfEmbeddedFontFamily family,
        PdfEmbeddedFontFallbackCandidate candidate) {
        byte[] selected = candidate.Style switch {
            OfficeFontStyle.Bold | OfficeFontStyle.Italic =>
                family.BoldItalicSnapshot
                ?? family.BoldSnapshot
                ?? family.ItalicSnapshot
                ?? family.RegularSnapshot,
            OfficeFontStyle.Bold =>
                family.BoldSnapshot ?? family.RegularSnapshot,
            OfficeFontStyle.Italic =>
                family.ItalicSnapshot ?? family.RegularSnapshot,
            _ => family.RegularSnapshot
        };
        return selected.SequenceEqual(candidate.DataSnapshot);
    }

    private string? FindReusablePromotedFallbackFamilyName(
        PdfEmbeddedFontFallbackCandidate candidate,
        PdfStandardFont slot,
        Dictionary<string, PdfEmbeddedFontFamily> prospectiveFamilies) {
        if (_embeddedFontFallbacks?.UsesNamedFontFamilies != true) {
            return null;
        }

        string stem = CreatePromotedFallbackFamilyNameStem(
            candidate.FontName,
            slot);
        foreach (PdfEmbeddedFontFallbackCandidate active in
            _embeddedFontFallbacks.Candidates) {
            if (!string.Equals(
                    active.PlannerFamilyName,
                    candidate.PlannerFamilyName,
                    StringComparison.OrdinalIgnoreCase)
                || !IsPromotedFallbackFamilyName(active.FontName, stem)
                || !prospectiveFamilies.ContainsKey(
                    NormalizeNamedFontFamilyKey(active.FontName))) {
                continue;
            }
            return active.FontName;
        }
        return null;
    }

    private static string CreatePromotedFallbackFamilyName(
        string originalName,
        PdfStandardFont slot,
        Dictionary<string, PdfEmbeddedFontFamily> prospectiveFamilies) {
        string stem = CreatePromotedFallbackFamilyNameStem(originalName, slot);
        string candidate = stem;
        int suffix = 2;
        while (prospectiveFamilies.ContainsKey(NormalizeNamedFontFamilyKey(candidate))) {
            candidate = $"{stem} {suffix++}";
        }
        return candidate;
    }

    private static string CreatePromotedFallbackFamilyNameStem(
        string originalName,
        PdfStandardFont slot) =>
        $"{originalName} [compatibility {PdfStandardFontMapper.GetFontFamily(slot)}]";

    private static bool IsPromotedFallbackFamilyName(
        string familyName,
        string stem) {
        if (string.Equals(familyName, stem, StringComparison.OrdinalIgnoreCase)) {
            return true;
        }
        string prefix = stem + " ";
        if (!familyName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }
        return int.TryParse(
                familyName.Remove(0, prefix.Length),
                System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture,
                out int suffix)
            && suffix >= 2;
    }

    private static PdfEmbeddedFontFallbackCandidate[] MergeFallbackCandidates(
        IEnumerable<PdfEmbeddedFontFallbackCandidate>? first,
        IEnumerable<PdfEmbeddedFontFallbackCandidate>? second) {
        var merged = new List<PdfEmbeddedFontFallbackCandidate>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (IEnumerable<PdfEmbeddedFontFallbackCandidate>? source in new[] {
            first,
            second
        }) {
            if (source == null) {
                continue;
            }
            foreach (PdfEmbeddedFontFallbackCandidate candidate in source) {
                if (names.Add(candidate.FontName)) {
                    merged.Add(candidate);
                }
            }
        }
        return merged.ToArray();
    }

    private static PdfEmbeddedFontFallbackCandidate[] OverlayFallbackCandidates(
        IEnumerable<PdfEmbeddedFontFallbackCandidate> existing,
        IEnumerable<PdfEmbeddedFontFallbackCandidate> overlay) {
        var replacements = overlay.ToDictionary(
            candidate => candidate.FontName,
            candidate => candidate,
            StringComparer.OrdinalIgnoreCase);
        var merged = new List<PdfEmbeddedFontFallbackCandidate>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (PdfEmbeddedFontFallbackCandidate candidate in existing) {
            PdfEmbeddedFontFallbackCandidate selected =
                replacements.TryGetValue(candidate.FontName, out PdfEmbeddedFontFallbackCandidate? replacement)
                    ? replacement
                    : candidate;
            if (names.Add(selected.FontName)) {
                merged.Add(selected);
            }
            replacements.Remove(candidate.FontName);
        }
        foreach (PdfEmbeddedFontFallbackCandidate candidate in overlay) {
            if (replacements.ContainsKey(candidate.FontName)
                && names.Add(candidate.FontName)) {
                merged.Add(candidate);
            }
        }
        return merged.ToArray();
    }

    private static PdfEmbeddedFontFallbackCandidate[] OverlayFallbackCandidateVariants(
        IEnumerable<PdfEmbeddedFontFallbackCandidate> existing,
        IEnumerable<PdfEmbeddedFontFallbackCandidate> overlay) {
        PdfEmbeddedFontFallbackCandidate[] overlayCandidates = overlay.ToArray();
        var replacements = overlayCandidates.ToDictionary(
            CandidateVariantKey,
            candidate => candidate,
            StringComparer.OrdinalIgnoreCase);
        var existingKeys = new HashSet<string>(
            existing.Select(CandidateVariantKey),
            StringComparer.OrdinalIgnoreCase);
        var merged = new List<PdfEmbeddedFontFallbackCandidate>();
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (PdfEmbeddedFontFallbackCandidate candidate in overlayCandidates) {
            string key = CandidateVariantKey(candidate);
            if (!existingKeys.Contains(key) && keys.Add(key)) {
                merged.Add(candidate);
            }
        }
        foreach (PdfEmbeddedFontFallbackCandidate candidate in existing) {
            string key = CandidateVariantKey(candidate);
            PdfEmbeddedFontFallbackCandidate selected =
                replacements.TryGetValue(key, out PdfEmbeddedFontFallbackCandidate? replacement)
                    ? replacement
                    : candidate;
            if (keys.Add(key)) {
                merged.Add(selected);
            }
            replacements.Remove(key);
        }
        return merged.ToArray();
    }

    private static PdfEmbeddedFontFallbackCandidate[] OverlayDeclaredFallbackCandidateVariants(
        IEnumerable<PdfEmbeddedFontFallbackCandidate> existing,
        IEnumerable<PdfEmbeddedFontFallbackCandidate> overlay) {
        PdfEmbeddedFontFallbackCandidate[] existingCandidates = existing.ToArray();
        PdfEmbeddedFontFallbackCandidate[] overlayCandidates = overlay.ToArray();
        string[] existingFamilies = existingCandidates
            .Select(CandidateFamilyKey)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        string[] overlayFamilies = overlayCandidates
            .Select(CandidateFamilyKey)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var merged = new List<PdfEmbeddedFontFallbackCandidate>();
        foreach (string family in existingFamilies) {
            PdfEmbeddedFontFallbackCandidate[] existingFamily = existingCandidates
                .Where(candidate => string.Equals(
                    CandidateFamilyKey(candidate),
                    family,
                    StringComparison.OrdinalIgnoreCase))
                .ToArray();
            PdfEmbeddedFontFallbackCandidate[] overlayFamily = overlayCandidates
                .Where(candidate => string.Equals(
                    CandidateFamilyKey(candidate),
                    family,
                    StringComparison.OrdinalIgnoreCase))
                .ToArray();
            merged.AddRange(overlayFamily.Length == 0
                ? existingFamily
                : OverlayFallbackCandidateVariants(existingFamily, overlayFamily));
        }
        foreach (string family in overlayFamilies.Where(family =>
            !existingFamilies.Contains(family, StringComparer.OrdinalIgnoreCase))) {
            merged.AddRange(overlayCandidates.Where(candidate => string.Equals(
                CandidateFamilyKey(candidate),
                family,
                StringComparison.OrdinalIgnoreCase)));
        }
        return merged.ToArray();
    }

    private static PdfEmbeddedFontFallbackCandidate[] ConcatDistinctCandidateVariants(
        IEnumerable<PdfEmbeddedFontFallbackCandidate> first,
        IEnumerable<PdfEmbeddedFontFallbackCandidate> second) {
        var merged = new List<PdfEmbeddedFontFallbackCandidate>();
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (PdfEmbeddedFontFallbackCandidate candidate in first.Concat(second)) {
            if (keys.Add(CandidateVariantKey(candidate))) {
                merged.Add(candidate);
            }
        }
        return merged.ToArray();
    }

    private static IEnumerable<string> EnumerateDeclaredFallbackFamilyNames(
        IEnumerable<PdfEmbeddedFontFallbackCandidate>? candidates) =>
        candidates?
            .Select(CandidateFamilyKey)
            .Distinct(StringComparer.OrdinalIgnoreCase)
        ?? Enumerable.Empty<string>();

    private static string CandidateFamilyKey(
        PdfEmbeddedFontFallbackCandidate candidate) =>
        string.IsNullOrWhiteSpace(candidate.PlannerFamilyName)
            ? candidate.FontName
            : candidate.PlannerFamilyName;

    private static string CandidateVariantKey(PdfEmbeddedFontFallbackCandidate candidate) =>
        candidate.FontName + "\u001f" + ((int)candidate.Style).ToString(
            System.Globalization.CultureInfo.InvariantCulture);

    private static PdfEmbeddedFontFallbackCandidate[] MergeProfileFallbackCandidates(
        PdfEmbeddedFontFallbackSet? existingFallbacks,
        IReadOnlyList<PdfEmbeddedFontFallbackCandidate> profileCandidates,
        HashSet<string> profileOwnedFallbackNames) {
        var merged = new List<PdfEmbeddedFontFallbackCandidate>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var replacements = new HashSet<string>(
            profileCandidates.Select(candidate => candidate.FontName),
            StringComparer.OrdinalIgnoreCase);
        if (existingFallbacks != null) {
            foreach (PdfEmbeddedFontFallbackCandidate candidate in existingFallbacks.Candidates) {
                if (profileOwnedFallbackNames.Contains(candidate.FontName)
                    && replacements.Contains(candidate.FontName)) {
                    continue;
                }
                if (names.Add(candidate.FontName)) {
                    merged.Add(candidate);
                }
            }
        }
        foreach (PdfEmbeddedFontFallbackCandidate candidate in profileCandidates) {
            if (names.Add(candidate.FontName)) {
                merged.Add(candidate);
            }
        }

        return merged.ToArray();
    }

    private void ClearEmbeddedStandardFontMappings() {
        _embeddedFonts?.Clear();
        _embeddedFontPrograms?.Clear();
        _embeddedOpenTypeCffFontPrograms?.Clear();
        _embeddedFontProgramFailures?.Clear();
        _usedEmbeddedFallbackFontSlots?.Clear();
    }

    private static Dictionary<string, PdfEmbeddedFontFallbackCandidate[]>?
        CloneRenderingProfileFamilyFallbacks(
            Dictionary<string, PdfEmbeddedFontFallbackCandidate[]>? source) {
        if (source == null) {
            return null;
        }

        var clone = new Dictionary<string, PdfEmbeddedFontFallbackCandidate[]>(
            StringComparer.OrdinalIgnoreCase);
        foreach (KeyValuePair<string, PdfEmbeddedFontFallbackCandidate[]> entry in source) {
            clone[entry.Key] = entry.Value
                .Select(candidate => new PdfEmbeddedFontFallbackCandidate(
                    candidate.FontName,
                    candidate.DataSnapshot,
                    candidate.UnicodeRanges,
                    candidate.Style,
                    candidate.PlannerFamilyName))
                .ToArray();
        }
        return clone;
    }

    private static ReadOnlyCollection<PdfEmbeddedFontFamily> CreateProfileFontFamilies(
        OfficeFontFaceCollection fonts) {
        var families = new List<PdfEmbeddedFontFamily>();
        foreach (IGrouping<string, OfficeFontFace> group in fonts.Faces
            .GroupBy(face => face.ResourceFamilyName, StringComparer.OrdinalIgnoreCase)) {
            OfficeFontFace[] faces = group.ToArray();
            OfficeFontFace? regular = SelectProfileFace(faces, OfficeFontStyle.Regular)
                ?? faces.FirstOrDefault();
            if (regular == null) {
                continue;
            }

            OfficeFontFace? bold = SelectProfileFace(faces, OfficeFontStyle.Bold);
            OfficeFontFace? italic = SelectProfileFace(faces, OfficeFontStyle.Italic);
            OfficeFontFace? boldItalic = SelectProfileFace(
                faces,
                OfficeFontStyle.Bold | OfficeFontStyle.Italic);
            families.Add(new PdfEmbeddedFontFamily(
                group.Key,
                regular.Data,
                bold?.Data,
                italic?.Data,
                boldItalic?.Data));
        }

        return families.AsReadOnly();
    }

    private static PdfEmbeddedFontFallbackCandidate[] CreateProfileFallbackCandidates(
        OfficeFontFaceCollection fonts) =>
        CreateProfileFallbackCandidates(fonts, fonts.FallbackFamilies);

    private static PdfEmbeddedFontFallbackCandidate[] CreateProfileFallbackCandidates(
        OfficeFontFaceCollection fonts,
        IEnumerable<string> fallbackFamilies) {
        var candidates = new List<PdfEmbeddedFontFallbackCandidate>();
        var addedVariants = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string fallbackFamily in fallbackFamilies) {
            OfficeFontFace[] matching = fonts.Faces
                .Where(face =>
                    string.Equals(face.FamilyName, fallbackFamily, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(face.ResourceFamilyName, fallbackFamily, StringComparison.OrdinalIgnoreCase))
                .Reverse()
                .ToArray();
            foreach (OfficeFontFace face in matching) {
                string key = face.ResourceFamilyName
                    + "\u001f"
                    + ((int)face.Style).ToString(
                        System.Globalization.CultureInfo.InvariantCulture);
                if (!addedVariants.Add(key)) {
                    continue;
                }

                candidates.Add(new PdfEmbeddedFontFallbackCandidate(
                    face.ResourceFamilyName,
                    face.Data,
                    face.UnicodeRanges,
                    face.Style,
                    string.Equals(
                        face.ResourceFamilyName,
                        fallbackFamily,
                        StringComparison.OrdinalIgnoreCase)
                        ? face.ResourceFamilyName
                        : face.FamilyName));
            }
        }

        return candidates.ToArray();
    }

    private static OfficeFontFace? SelectProfileFace(
        IEnumerable<OfficeFontFace> faces,
        OfficeFontStyle style) {
        OfficeFontStyle normalized = style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic);
        return faces.FirstOrDefault(face =>
            (face.Style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic)) == normalized);
    }
}
