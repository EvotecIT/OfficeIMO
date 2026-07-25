using OfficeIMO.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static PdfCore.PdfStandardFont MapFont(
        string familyName,
        IReadOnlyDictionary<string, PdfCore.PdfStandardFont> webFonts) {
        foreach (string candidate in EnumerateFamilies(familyName)) {
            if (webFonts.TryGetValue(candidate, out PdfCore.PdfStandardFont embedded)) {
                return embedded;
            }
        }

        return MapStandardFont(familyName);
    }

    private static PdfCore.PdfStandardFont MapStandardFont(string familyName) {
        return PdfCore.PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfCore.PdfStandardFont font)
            ? font
            : PdfCore.PdfStandardFont.Helvetica;
    }

    private static IReadOnlyDictionary<string, PdfCore.PdfStandardFont> RegisterWebFonts(
        PdfCore.PdfDocument pdf,
        HtmlRenderDocument rendered,
        HtmlDiagnosticReport? diagnostics,
        ISet<PdfCore.PdfStandardFont> reservedFontSlots,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeFontFaceCollection faces = rendered.Fonts;
        var byFamily = faces.Faces
            .GroupBy(face => face.ResourceFamilyName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
        var mappings = new Dictionary<string, PdfCore.PdfStandardFont>(StringComparer.OrdinalIgnoreCase);
        if (byFamily.Count == 0) return mappings;

        var orderedFamilies = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string familyNames in EnumerateUsedFontFamilyLists(rendered.Pages.SelectMany(page => page.Visuals))) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (string family in EnumerateFamilies(familyNames)) {
                if (byFamily.ContainsKey(family) && seen.Add(family)) orderedFamilies.Add(family);
            }
        }

        foreach (string family in orderedFamilies) {
            cancellationToken.ThrowIfCancellationRequested();
            RegisterNamedFamily(pdf, family, byFamily[family], cancellationToken);
            mappings[family] = MapStandardFont(family);
        }

        return mappings;
    }

    private static void ReserveUsedStandardFontSlots(
        HtmlRenderDocument rendered,
        ISet<string> activeWebFontFamilies,
        ISet<PdfCore.PdfStandardFont> reservedFontSlots) {
        foreach (string familyNames in EnumerateUsedFontFamilyLists(rendered.Pages.SelectMany(page => page.Visuals))) {
            if (EnumerateFamilies(familyNames).Any(activeWebFontFamilies.Contains)) continue;
            reservedFontSlots.Add(PdfCore.PdfStandardFontMapper.GetFontFamily(MapStandardFont(familyNames)));
        }
    }

    private static void RegisterUsedSystemFontFamilies(
        PdfCore.PdfDocument pdf,
        HtmlRenderDocument rendered,
        ISet<string> activeWebFontFamilies,
        ISet<PdfCore.PdfStandardFont> reservedFontSlots,
        CancellationToken cancellationToken) {
        List<HtmlRenderText> textRuns = EnumerateVisuals(rendered.Pages.SelectMany(page => page.Visuals))
            .OfType<HtmlRenderText>()
            .Where(text => !EnumerateFamilies(text.Font.FamilyName).Any(activeWebFontFamilies.Contains))
            .ToList();
        int loadedFamilyCount = 0;

        foreach (string familyName in textRuns
                     .SelectMany(text => EnumerateFamilies(text.Font.FamilyName))
                     .Distinct(StringComparer.OrdinalIgnoreCase)
                     .Take(MaximumSystemFontFamilyCandidates)) {
            cancellationToken.ThrowIfCancellationRequested();
            List<HtmlRenderText> familyRuns = textRuns
                .Where(text => EnumerateFamilies(text.Font.FamilyName).Contains(familyName, StringComparer.OrdinalIgnoreCase))
                .ToList();
            if (familyRuns.Count == 0 || pdf.Options.HasNamedFontFamily(familyName)) continue;
            if (loadedFamilyCount >= MaximumLoadedSystemFontFamilies) break;
            if (!PdfCore.PdfEmbeddedFontFamily.TryFromSystem(familyName, out PdfCore.PdfEmbeddedFontFamily? family)
                || family == null) continue;

            pdf.Options.RegisterNamedFontFamily(CreateCoverageSafeFontFamily(family, familyRuns));
            reservedFontSlots.Add(PdfCore.PdfStandardFontMapper.GetFontFamily(MapStandardFont(familyName)));
            loadedFamilyCount++;
        }
    }

    private static PdfCore.PdfEmbeddedFontFamily CreateCoverageSafeFontFamily(
        PdfCore.PdfEmbeddedFontFamily family,
        IEnumerable<HtmlRenderText> textRuns) {
        List<HtmlRenderText> runs = textRuns.ToList();
        byte[] regular = family.Regular;
        byte[]? bold = SelectCoverageSafeFace(
            family.Bold,
            regular,
            runs.Where(run => run.Font.IsBold && !run.Font.IsItalic).Select(run => run.Text));
        byte[]? italic = SelectCoverageSafeFace(
            family.Italic,
            regular,
            runs.Where(run => !run.Font.IsBold && run.Font.IsItalic).Select(run => run.Text));
        byte[]? boldItalic = SelectCoverageSafeFace(
            family.BoldItalic ?? family.Bold ?? family.Italic,
            regular,
            runs.Where(run => run.Font.IsBold && run.Font.IsItalic).Select(run => run.Text));
        return new PdfCore.PdfEmbeddedFontFamily(family.FamilyName, regular, bold, italic, boldItalic);
    }

    private static byte[]? SelectCoverageSafeFace(
        byte[]? styledFace,
        byte[] regularFace,
        IEnumerable<string> requiredText) {
        if (styledFace == null) return null;
        string text = string.Concat(requiredText);
        if (text.Length == 0 || FontCoversText(styledFace, text) || !FontCoversText(regularFace, text)) {
            return styledFace;
        }
        return regularFace;
    }

    private static bool FontCoversText(byte[] fontData, string text) {
        var candidate = new PdfCore.PdfEmbeddedFontFallbackCandidate("HTML system font coverage", fontData);
        return PdfCore.PdfTextDiagnostics.PlanEmbeddedFontFallbackText(text, new[] { candidate }).IsFullyCovered;
    }

    private static IEnumerable<string> EnumerateUsedFontFamilyLists(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in EnumerateVisuals(visuals)) {
            if (visual is HtmlRenderText text) {
                yield return text.Font.FamilyName;
            } else if (visual is HtmlRenderDrawing drawing) {
                foreach (string familyNames in EnumerateDrawingFontFamilyLists(drawing.Drawing.Elements)) {
                    yield return familyNames;
                }
            }
        }
    }

    private static IEnumerable<string> EnumerateDrawingFontFamilyLists(IEnumerable<OfficeDrawingElement> elements) {
        foreach (OfficeDrawingElement element in elements) {
            if (element is OfficeDrawingText text) {
                yield return text.Font.FamilyName;
            } else if (element is OfficeDrawingEffectGroup effectGroup) {
                foreach (string familyNames in EnumerateDrawingFontFamilyLists(effectGroup.Drawing.Elements)) {
                    yield return familyNames;
                }
            } else if (element is OfficeDrawingTilingPattern tilingPattern) {
                foreach (string familyNames in EnumerateDrawingFontFamilyLists(tilingPattern.Tile.Elements)) {
                    yield return familyNames;
                }
            }
        }
    }

    private static IEnumerable<HtmlRenderVisual> EnumerateVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual is HtmlRenderClipGroup clipGroup
                ? clipGroup.Visuals
                : visual is HtmlRenderPathClipGroup pathClipGroup
                    ? pathClipGroup.Visuals
                    : visual is HtmlRenderEffectGroup effectGroup
                        ? effectGroup.Visuals
                        : visual is HtmlRenderSemanticGroup semanticGroup
                            ? semanticGroup.Visuals
                            : visual is HtmlRenderLogicalTextGroup logicalTextGroup
                                ? logicalTextGroup.Visuals
                                : null;
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateVisuals(children)) yield return child;
        }
    }

    internal static PdfCore.PdfTextFallbackFeatures ResolveTextFallbackFeatures(
        HtmlRenderDocument rendered,
        PdfCore.PdfTextFallbackFeatures requested) {
        if (requested == PdfCore.PdfTextFallbackFeatures.None) return requested;

        foreach (HtmlRenderVisual visual in EnumerateVisuals(rendered.Pages.SelectMany(page => page.Visuals))) {
            if (visual is HtmlRenderText text && RequiresUnicodeFont(text.Text)) return requested;
            if (visual is HtmlRenderDrawing drawing && DrawingRequiresUnicodeFont(drawing.Drawing.Elements)) {
                return requested;
            }
        }

        return PdfCore.PdfTextFallbackFeatures.None;
    }

    private static bool DrawingRequiresUnicodeFont(IEnumerable<OfficeDrawingElement> elements) {
        foreach (OfficeDrawingElement element in elements) {
            if (element is OfficeDrawingText text && RequiresUnicodeFont(text.Text)) return true;
            if (element is OfficeDrawingEffectGroup effectGroup
                && DrawingRequiresUnicodeFont(effectGroup.Drawing.Elements)) return true;
            if (element is OfficeDrawingTilingPattern tilingPattern
                && DrawingRequiresUnicodeFont(tilingPattern.Tile.Elements)) return true;
        }

        return false;
    }

    private static bool RequiresUnicodeFont(string text) =>
        PdfCore.PdfTextDiagnostics.AnalyzeWinAnsiText(text).Count != 0;

    private static void RegisterNamedFamily(
        PdfCore.PdfDocument pdf,
        string family,
        IReadOnlyList<OfficeFontFace> faces,
        CancellationToken cancellationToken) {
        OfficeFontFace regular = FindFace(faces, OfficeFontStyle.Regular) ?? faces[0];
        OfficeFontFace bold = FindFace(faces, OfficeFontStyle.Bold) ?? regular;
        OfficeFontFace italic = FindFace(faces, OfficeFontStyle.Italic) ?? regular;
        OfficeFontFace boldItalic = FindFace(faces, OfficeFontStyle.Bold | OfficeFontStyle.Italic) ?? bold;
        cancellationToken.ThrowIfCancellationRequested();
        pdf.Options.RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily(
            family,
            regular.Data,
            bold.Data,
            italic.Data,
            boldItalic.Data));
    }

    private static OfficeFontFace? FindFace(IReadOnlyList<OfficeFontFace> faces, OfficeFontStyle style) {
        OfficeFontStyle normalized = style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic);
        return faces.FirstOrDefault(face =>
            (face.Style & (OfficeFontStyle.Bold | OfficeFontStyle.Italic)) == normalized);
    }

    private static IEnumerable<string> EnumerateFamilies(string? familyNames) {
        if (string.IsNullOrWhiteSpace(familyNames)) yield break;

        foreach (string raw in familyNames!.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)) {
            string family = raw.Trim().Trim('"', '\'');
            if (family.Length > 0) yield return family;
        }
    }
}
