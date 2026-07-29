using OfficeIMO.Drawing;
using System.Collections.ObjectModel;

namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
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
        ReadOnlyCollection<PdfEmbeddedFontFamily> families = CreateProfileFontFamilies(profileFonts);
        if (mode == OfficeRenderingProfileApplyMode.Replace) {
            ClearNamedFontFamilies();
            EmbeddedFontFallbacks = null;
        }
        if (families.Count > 0) {
            foreach (PdfEmbeddedFontFamily family in families) {
                RegisterNamedFontFamily(family);
            }

            if (mode == OfficeRenderingProfileApplyMode.Replace
                || EmbeddedFontFallbacksSnapshot == null) {
                PdfEmbeddedFontFallbackCandidate[] candidates =
                    CreateProfileFallbackCandidates(profileFonts);
                if (candidates.Length > 0) {
                    RegisterEmbeddedFontFallbacks(new PdfEmbeddedFontFallbackSet(candidates));
                }
            }
        }

        return this;
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
