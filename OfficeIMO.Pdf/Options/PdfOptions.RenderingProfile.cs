using OfficeIMO.Drawing;
using System.Collections.ObjectModel;

namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    private Dictionary<string, PdfEmbeddedFontFallbackCandidate[]>? _renderingProfileFamilyFallbacks;
    private PdfEmbeddedFontFallbackCandidate[]? _renderingProfileDeclaredFallbackCandidates;

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

        if (mode == OfficeRenderingProfileApplyMode.Replace || profile.TextShapingProvider != null) {
            TextShapingProvider = profile.TextShapingProvider;
        }
        if (mode == OfficeRenderingProfileApplyMode.Replace || profile.TextShapingLanguage != null) {
            Language = profile.TextShapingLanguage;
        }

        OfficeFontFaceCollection profileFonts = profile.Fonts;
        PdfEmbeddedFontFallbackSet? existingFallbacks = mode == OfficeRenderingProfileApplyMode.Overlay
            ? EmbeddedFontFallbacksSnapshot?.Clone()
            : null;
        var preservedFallbackNames = new HashSet<string>(
            existingFallbacks?.Candidates.Select(candidate => candidate.FontName)
                ?? Enumerable.Empty<string>(),
            StringComparer.OrdinalIgnoreCase);
        var preservedNamedFallbackNames = new HashSet<string>(
            existingFallbacks?.UsesNamedFontFamilies == true
                ? existingFallbacks.FontFamilyNames
                : Array.Empty<string>(),
            StringComparer.OrdinalIgnoreCase);
        ReadOnlyCollection<PdfEmbeddedFontFamily> families = CreateProfileFontFamilies(profileFonts);
        if (mode == OfficeRenderingProfileApplyMode.Replace) {
            ClearNamedFontFamilies();
            ClearEmbeddedStandardFontMappings();
            _embeddedFontFallbacks = null;
            _renderingProfileFamilyFallbacks?.Clear();
            _renderingProfileDeclaredFallbackCandidates = null;
        }
        if (families.Count > 0) {
            foreach (PdfEmbeddedFontFamily family in families) {
                if (mode == OfficeRenderingProfileApplyMode.Overlay
                    && preservedNamedFallbackNames.Contains(family.FamilyName)) {
                    continue;
                }
                RegisterNamedFontFamily(family);
            }
        }

        RegisterProfileFamilyFallbacks(profileFonts);

        PdfEmbeddedFontFallbackCandidate[] profileCandidates =
            CreateProfileFallbackCandidates(profileFonts);
        PdfEmbeddedFontFallbackCandidate[] profileCandidateVariants = mode == OfficeRenderingProfileApplyMode.Overlay
            ? profileCandidates
                .Where(candidate => !preservedFallbackNames.Contains(candidate.FontName))
                .ToArray()
            : profileCandidates;
        _renderingProfileDeclaredFallbackCandidates =
            _renderingProfileDeclaredFallbackCandidates == null
                ? profileCandidateVariants
                : OverlayFallbackCandidateVariants(
                    _renderingProfileDeclaredFallbackCandidates,
                    profileCandidateVariants);
        PdfEmbeddedFontFallbackCandidate[] regularProfileCandidates =
            SelectRenderingProfileCandidates(
                profileCandidateVariants,
                bold: false,
                italic: false);
        PdfEmbeddedFontFallbackCandidate[] combinedCandidates =
            MergeProfileFallbackCandidates(existingFallbacks, regularProfileCandidates);
        if (combinedCandidates.Length > 0) {
            foreach (PdfEmbeddedFontFallbackCandidate candidate in combinedCandidates) {
                if (!HasNamedFontFamily(candidate.FontName)) {
                    RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                        candidate.FontName,
                        candidate.DataSnapshot));
                }
            }

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

        foreach (string familyCandidate in
            EnumerateOfficeFontFamilyCandidates(familyName!)) {
            if (_renderingProfileFamilyFallbacks.TryGetValue(
                    familyCandidate,
                    out candidates)) {
                return true;
            }
        }
        return false;
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
        return combinedCandidates.Length == 0
            ? null
            : new PdfEmbeddedFontFallbackSet(combinedCandidates);
    }

    private void RegisterProfileFamilyFallbacks(OfficeFontFaceCollection fonts) {
        var scopedFamilies = new HashSet<string>(
            fonts.Faces
                .Where(face => !face.UnicodeRanges.IsAll)
                .Select(face => face.FamilyName),
            StringComparer.OrdinalIgnoreCase);
        foreach (IGrouping<string, OfficeFontFace> family in fonts.Faces
            .Where(face => scopedFamilies.Contains(face.FamilyName))
            .GroupBy(face => face.FamilyName, StringComparer.OrdinalIgnoreCase)) {
            PdfEmbeddedFontFallbackCandidate[] candidates = family
                .Where(face => !face.UnicodeRanges.IsAll)
                .Reverse()
                .Concat(family
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
                    face.Style))
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
        OfficeFontStyle[] precedence = RenderingProfileStylePrecedence(bold, italic);
        var selected = new List<PdfEmbeddedFontFallbackCandidate>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (OfficeFontStyle style in precedence) {
            foreach (PdfEmbeddedFontFallbackCandidate candidate in candidates) {
                if (candidate.Style == style && names.Add(candidate.FontName)) {
                    selected.Add(candidate);
                }
            }
        }
        return selected.ToArray();
    }

    private static PdfEmbeddedFontFallbackCandidate[] SelectRenderingProfileCandidates(
        IReadOnlyList<PdfEmbeddedFontFallbackCandidate> candidates,
        bool bold,
        bool italic) {
        OfficeFontStyle[] precedence = RenderingProfileStylePrecedence(bold, italic);
        var selected = new List<PdfEmbeddedFontFallbackCandidate>();
        foreach (IGrouping<string, PdfEmbeddedFontFallbackCandidate> family in candidates
            .GroupBy(candidate => candidate.FontName, StringComparer.OrdinalIgnoreCase)) {
            foreach (OfficeFontStyle style in precedence) {
                PdfEmbeddedFontFallbackCandidate? candidate =
                    family.FirstOrDefault(item => item.Style == style);
                if (candidate != null) {
                    selected.Add(candidate);
                    break;
                }
            }
        }
        return selected.ToArray();
    }

    private static OfficeFontStyle[] RenderingProfileStylePrecedence(
        bool bold,
        bool italic) {
        OfficeFontStyle requested =
            (bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        return requested switch {
            OfficeFontStyle.Bold | OfficeFontStyle.Italic => new[] {
                OfficeFontStyle.Bold | OfficeFontStyle.Italic,
                OfficeFontStyle.Regular,
                OfficeFontStyle.Bold,
                OfficeFontStyle.Italic
            },
            OfficeFontStyle.Bold => new[] {
                OfficeFontStyle.Bold,
                OfficeFontStyle.Regular
            },
            OfficeFontStyle.Italic => new[] {
                OfficeFontStyle.Italic,
                OfficeFontStyle.Regular
            },
            _ => new[] { OfficeFontStyle.Regular }
        };
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
        var replacements = overlay.ToDictionary(
            CandidateVariantKey,
            candidate => candidate,
            StringComparer.OrdinalIgnoreCase);
        var merged = new List<PdfEmbeddedFontFallbackCandidate>();
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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
        foreach (PdfEmbeddedFontFallbackCandidate candidate in overlay) {
            string key = CandidateVariantKey(candidate);
            if (replacements.ContainsKey(key) && keys.Add(key)) {
                merged.Add(candidate);
            }
        }
        return merged.ToArray();
    }

    private static string CandidateVariantKey(PdfEmbeddedFontFallbackCandidate candidate) =>
        candidate.FontName + "\u001f" + ((int)candidate.Style).ToString(
            System.Globalization.CultureInfo.InvariantCulture);

    private static PdfEmbeddedFontFallbackCandidate[] MergeProfileFallbackCandidates(
        PdfEmbeddedFontFallbackSet? existingFallbacks,
        IReadOnlyList<PdfEmbeddedFontFallbackCandidate> profileCandidates) {
        var merged = new List<PdfEmbeddedFontFallbackCandidate>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (existingFallbacks != null) {
            foreach (PdfEmbeddedFontFallbackCandidate candidate in existingFallbacks.Candidates) {
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
                    candidate.Style))
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
        OfficeFontFaceCollection fonts) {
        var candidates = new List<PdfEmbeddedFontFallbackCandidate>();
        var addedVariants = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string fallbackFamily in fonts.FallbackFamilies) {
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
                    face.Style));
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
