using OfficeIMO.Drawing;
using System.Collections.ObjectModel;

namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    private Dictionary<string, PdfEmbeddedFontFallbackSet>? _renderingProfileFamilyFallbacks;

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
        ReadOnlyCollection<PdfEmbeddedFontFamily> families = CreateProfileFontFamilies(profileFonts);
        if (mode == OfficeRenderingProfileApplyMode.Replace) {
            ClearNamedFontFamilies();
            ClearEmbeddedStandardFontMappings();
            _embeddedFontFallbacks = null;
            _renderingProfileFamilyFallbacks?.Clear();
        }
        if (families.Count > 0) {
            foreach (PdfEmbeddedFontFamily family in families) {
                if (mode == OfficeRenderingProfileApplyMode.Overlay
                    && preservedFallbackNames.Contains(family.FamilyName)) {
                    continue;
                }
                RegisterNamedFontFamily(family);
            }
        }

        RegisterProfileFamilyFallbacks(profileFonts);

        PdfEmbeddedFontFallbackCandidate[] profileCandidates =
            CreateProfileFallbackCandidates(profileFonts);
        PdfEmbeddedFontFallbackCandidate[] combinedCandidates =
            MergeProfileFallbackCandidates(existingFallbacks, profileCandidates);
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
        if (string.IsNullOrWhiteSpace(familyName)
            || _renderingProfileFamilyFallbacks == null
            || !_renderingProfileFamilyFallbacks.TryGetValue(
                familyName!.Trim(),
                out PdfEmbeddedFontFallbackSet? registered)) {
            return false;
        }

        fallbackSet = registered;
        return true;
    }

    internal bool TryGetEffectiveRenderingProfileFallbacks(
        string? familyName,
        bool bold,
        bool italic,
        out PdfEmbeddedFontFallbackSet? fallbackSet) {
        fallbackSet = null;
        if (!TryGetRenderingProfileFamilyFallbacks(
                familyName,
                out PdfEmbeddedFontFallbackSet? registered)
            || registered == null) {
            return false;
        }

        PdfEmbeddedFontFallbackCandidate[] styledCandidates =
            SelectRenderingProfileCandidates(registered.Candidates, bold, italic);
        PdfEmbeddedFontFallbackCandidate[] combinedCandidates =
            MergeFallbackCandidates(styledCandidates, _embeddedFontFallbacks?.Candidates);
        if (combinedCandidates.Length == 0) {
            return false;
        }

        fallbackSet = new PdfEmbeddedFontFallbackSet(combinedCandidates);
        return true;
    }

    private void RegisterProfileFamilyFallbacks(OfficeFontFaceCollection fonts) {
        foreach (IGrouping<string, OfficeFontFace> family in fonts.Faces
            .Where(face => !face.UnicodeRanges.IsAll)
            .GroupBy(face => face.FamilyName, StringComparer.OrdinalIgnoreCase)) {
            PdfEmbeddedFontFallbackCandidate[] candidates = family
                .GroupBy(face => face.ResourceFamilyName, StringComparer.OrdinalIgnoreCase)
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

            Dictionary<string, PdfEmbeddedFontFallbackSet> fallbacks =
                _renderingProfileFamilyFallbacks ??=
                    new Dictionary<string, PdfEmbeddedFontFallbackSet>(
                        StringComparer.OrdinalIgnoreCase);
            PdfEmbeddedFontFallbackCandidate[] merged = fallbacks.TryGetValue(
                    family.Key,
                    out PdfEmbeddedFontFallbackSet? existing)
                ? MergeFallbackCandidates(existing.Candidates, candidates)
                : candidates;
            fallbacks[family.Key] = new PdfEmbeddedFontFallbackSet(merged);
        }
    }

    private static PdfEmbeddedFontFallbackCandidate[] SelectRenderingProfileCandidates(
        IReadOnlyList<PdfEmbeddedFontFallbackCandidate> candidates,
        bool bold,
        bool italic) {
        OfficeFontStyle requested =
            (bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        OfficeFontStyle[] precedence = requested switch {
            OfficeFontStyle.Bold | OfficeFontStyle.Italic => new[] {
                OfficeFontStyle.Bold | OfficeFontStyle.Italic,
                OfficeFontStyle.Bold,
                OfficeFontStyle.Italic,
                OfficeFontStyle.Regular
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

    private static Dictionary<string, PdfEmbeddedFontFallbackSet>?
        CloneRenderingProfileFamilyFallbacks(
            Dictionary<string, PdfEmbeddedFontFallbackSet>? source) {
        if (source == null) {
            return null;
        }

        var clone = new Dictionary<string, PdfEmbeddedFontFallbackSet>(
            StringComparer.OrdinalIgnoreCase);
        foreach (KeyValuePair<string, PdfEmbeddedFontFallbackSet> entry in source) {
            clone[entry.Key] = entry.Value.Clone();
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
        var addedResources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string fallbackFamily in fonts.FallbackFamilies) {
            OfficeFontFace[] matching = fonts.Faces
                .Where(face =>
                    string.Equals(face.FamilyName, fallbackFamily, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(face.ResourceFamilyName, fallbackFamily, StringComparison.OrdinalIgnoreCase))
                .Reverse()
                .ToArray();
            foreach (IGrouping<string, OfficeFontFace> group in matching
                .GroupBy(face => face.ResourceFamilyName, StringComparer.OrdinalIgnoreCase)) {
                if (!addedResources.Add(group.Key)) {
                    continue;
                }

                OfficeFontFace[] faces = group.ToArray();
                OfficeFontFace? regular = SelectProfileFace(faces, OfficeFontStyle.Regular)
                    ?? faces.FirstOrDefault();
                if (regular != null) {
                    candidates.Add(new PdfEmbeddedFontFallbackCandidate(
                        regular.ResourceFamilyName,
                        regular.Data,
                        regular.UnicodeRanges));
                }
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
