using OfficeIMO.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private static Func<string, OfficeFontInfo, OfficeFontFaceDescriptor, HtmlTextFaceMetrics?>? CreateFallbackTextFaceMetrics(
        HtmlToPdfOptions options) {
        if (!options.ResourcePolicy.AllowSystemFontEmbedding
            || !options.ResourcePolicy.AllowDocumentFontEmbedding) return null;
        return (text, font, descriptor) => {
            foreach (string family in EnumerateBoundedSystemFamilies(font.FamilyName)) {
                if (PdfCore.PdfEmbeddedFontFamily.TryMeasureSystemFaceVerticalMetrics(
                    family, descriptor, text, font.Size, out double height, out double baseline)) {
                    return new HtmlTextFaceMetrics(height, baseline);
                }
            }
            return null;
        };
    }

    private static Func<string, OfficeFontInfo, OfficeFontFaceDescriptor, double?> CreateFallbackTextMeasurement(
        HtmlToPdfOptions options,
        PdfCore.PdfOptions measurementOptions) {
        bool useInstalledFonts = options.ResourcePolicy.AllowSystemFontEmbedding
            && options.ResourcePolicy.AllowDocumentFontEmbedding;
        var attemptedFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int loadedFamilies = 0;
        return (text, font, descriptor) => {
            string measuredFamily = font.FamilyName;
            if (useInstalledFonts && font.FamilyName.Length > 0
                && HtmlRenderCssValues.TrySplitTopLevelCommas(
                    font.FamilyName, MaximumCssFontFamilyCandidates, out _)) {
                IReadOnlyList<string> families = EnumerateFamilies(font.FamilyName).ToList();
                foreach (string familyName in families) {
                    if (NeedsNumericSystemFace(descriptor)
                        && TryRegisterSystemFace(measurementOptions, familyName, text, descriptor,
                            ref loadedFamilies, attemptedFamilies, out string? selectedFamily)
                        && NamedFontCoversText(measurementOptions, selectedFamily!, text,
                            font.IsBold, font.IsItalic)) {
                        measuredFamily = selectedFamily!;
                        break;
                    }
                    if (!measurementOptions.HasNamedFontFamily(familyName)
                        && loadedFamilies < MaximumLoadedSystemFontFamilies
                        && attemptedFamilies.Count < MaximumSystemFontFamilyCandidates
                        && familyName.Length <= 256
                        && attemptedFamilies.Add(familyName)
                        && PdfCore.PdfEmbeddedFontFamily.TryFromSystem(
                            familyName, out PdfCore.PdfEmbeddedFontFamily? family)
                        && family != null
                        && measurementOptions.TryRegisterNamedFontFamily(
                            CreateCoverageStableFontFamily(family))) {
                        loadedFamilies++;
                    }
                    if (NamedFontCoversText(measurementOptions, familyName, text, font.IsBold, font.IsItalic)) {
                        measuredFamily = familyName;
                        break;
                    }
                }
            }
            return PdfCore.PdfWriter.MeasurePositionedText(
                new PdfCore.PdfTextRun(text, bold: font.IsBold, italic: font.IsItalic,
                    fontSize: font.Size, font: MapStandardFont(measuredFamily), fontFamily: measuredFamily),
                measurementOptions);
        };
    }

    private static bool NeedsNumericSystemFace(OfficeFontFaceDescriptor descriptor) =>
        descriptor.Weight != (descriptor.Weight >= 600 ? 700 : 400)
        || descriptor.StretchPercent != 100D
        || descriptor.Slant == OfficeFontSlant.Oblique;

    private static bool TryRegisterSystemFace(
        PdfCore.PdfOptions options,
        string familyName,
        string text,
        OfficeFontFaceDescriptor descriptor,
        ref int loadedFamilies,
        ISet<string> attemptedFamilies,
        out string? selectedFamily) {
        selectedFamily = null;
        if (familyName.Length > 256
            || !PdfCore.PdfEmbeddedFontFamily.TryResolveSystemFace(
                familyName, descriptor, text, out PdfCore.PdfEmbeddedFontFamily? face)
            || face == null) return false;
        selectedFamily = face.FamilyName;
        if (options.HasNamedFontFamily(selectedFamily)) return true;
        if (loadedFamilies >= MaximumLoadedSystemFontFamilies
            || attemptedFamilies.Count >= MaximumSystemFontFamilyCandidates
            || !attemptedFamilies.Add(selectedFamily)) return false;
        if (!options.TryRegisterNamedFontFamily(face)) return false;
        loadedFamilies++;
        return true;
    }

    private static string ResolvePdfFontFamilyForText(
        string familyNames,
        string text,
        bool bold,
        bool italic,
        OfficeFontFaceDescriptor descriptor,
        bool allowInstalledFaces,
        PdfCore.PdfOptions options) {
        if (allowInstalledFaces && NeedsNumericSystemFace(descriptor)) {
            foreach (string familyName in EnumerateBoundedSystemFamilies(familyNames)) {
                if (PdfCore.PdfEmbeddedFontFamily.TryResolveSystemFace(
                        familyName, descriptor, text, out PdfCore.PdfEmbeddedFontFamily? face)
                    && face != null
                    && NamedFontCoversText(options, face.FamilyName, text, bold, italic)) {
                    return face.FamilyName;
                }
            }
        }
        if (familyNames.IndexOf(',') < 0) return familyNames;
        foreach (string familyName in EnumerateBoundedSystemFamilies(familyNames)) {
            if (NamedFontCoversText(options, familyName, text, bold, italic)) return familyName;
        }
        return familyNames;
    }

    private static bool NamedFontCoversText(
        PdfCore.PdfOptions options,
        string familyName,
        string text,
        bool bold,
        bool italic) {
        if (!options.TryResolveNamedFontFace(familyName, bold, italic, out PdfCore.PdfNamedFontFace face)) {
            return false;
        }
        if (options.TryGetNamedFontProgram(face, out PdfCore.PdfTrueTypeFontProgram? trueType)
            && trueType != null
            && PdfCore.PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, trueType).Count == 0) {
            return true;
        }
        return options.TryGetNamedOpenTypeCffFontProgram(face, out PdfCore.PdfOpenTypeCffFontProgram? cff)
            && cff != null
            && PdfCore.PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, cff).Count == 0;
    }

    private static PdfCore.PdfStandardFont MapFont(
        string familyName,
        string text,
        OfficeFontStyle style,
        RegisteredWebFonts webFonts) {
        IReadOnlyList<OfficeFontFallbackRun> planned = webFonts.Faces.PlanFallbackRuns(text, familyName, style);
        if (planned.Count == 1
            && string.Equals(planned[0].Text, text, StringComparison.Ordinal)
            && webFonts.Slots.TryGetValue(planned[0].FamilyName, out PdfCore.PdfStandardFont embedded)) {
            return embedded;
        }

        return MapStandardFont(familyName);
    }

    private static PdfCore.PdfStandardFont MapStandardFont(string familyName) {
        return PdfCore.PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfCore.PdfStandardFont font)
            ? font
            : PdfCore.PdfStandardFont.Helvetica;
    }

    private static RegisteredWebFonts RegisterWebFonts(
        PdfCore.PdfDocument pdf,
        HtmlRenderDocument rendered,
        HtmlDiagnosticReport diagnostics,
        int maxOutlinedTextCharactersPerRun,
        int maxOutlinedTextPathCommands,
        IOfficeTextShapingProvider? textShapingProvider,
        string? textShapingLanguage,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeFontFaceCollection faces = rendered.Fonts;
        var byFamily = faces.Faces
            .Where(face => face.CanEmbedAsStaticPdfFont)
            .GroupBy(face => face.ResourceFamilyName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
        var mappings = new Dictionary<string, PdfCore.PdfStandardFont>(StringComparer.OrdinalIgnoreCase);
        var outlineBudget = new OutlinedTextBudget(
            maxOutlinedTextCharactersPerRun,
            maxOutlinedTextPathCommands);
        if (byFamily.Count == 0) return new RegisteredWebFonts(
            mappings,
            faces,
            pdf.Options,
            diagnostics,
            outlineBudget,
            textShapingProvider,
            textShapingLanguage);

        var orderedFamilies = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string familyNames in EnumerateUsedWebFontFamilyLists(
                     rendered.Pages.SelectMany(page => page.Visuals),
                     faces)) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (string family in EnumerateFamilies(familyNames)) {
                if (byFamily.ContainsKey(family) && seen.Add(family)) orderedFamilies.Add(family);
            }
        }

        foreach (string family in orderedFamilies) {
            cancellationToken.ThrowIfCancellationRequested();
            if (RegisterNamedFamily(pdf, family, byFamily[family], cancellationToken)) {
                mappings[family] = MapStandardFont(family);
            }
        }

        return new RegisteredWebFonts(
            mappings,
            faces,
            pdf.Options,
            diagnostics,
            outlineBudget,
            textShapingProvider,
            textShapingLanguage);
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
            .ToList();
        int loadedFamilyCount = 0;
        var attemptedFaces = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (HtmlRenderText run in textRuns) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!NeedsNumericSystemFace(run.FontDescriptor)) continue;
            foreach (string familyName in EnumerateBoundedSystemFamilies(run.Font.FamilyName)) {
                if (activeWebFontFamilies.Contains(familyName)) continue;
                if (!TryRegisterSystemFace(pdf.Options, familyName, run.Text, run.FontDescriptor,
                        ref loadedFamilyCount, attemptedFaces, out string? selectedFamily)) continue;
                reservedFontSlots.Add(PdfCore.PdfStandardFontMapper.GetFontFamily(MapStandardFont(familyName)));
                if (NamedFontCoversText(pdf.Options, selectedFamily!, run.Text,
                        run.Font.IsBold, run.Font.IsItalic)) break;
            }
            if (loadedFamilyCount >= MaximumLoadedSystemFontFamilies) break;
        }

        foreach (string familyName in textRuns
                     .SelectMany(text => EnumerateBoundedSystemFamilies(text.Font.FamilyName))
                     .Distinct(StringComparer.OrdinalIgnoreCase)
                     .Take(MaximumSystemFontFamilyCandidates)) {
            cancellationToken.ThrowIfCancellationRequested();
            List<HtmlRenderText> familyRuns = textRuns
                .Where(text => EnumerateBoundedSystemFamilies(text.Font.FamilyName).Contains(familyName, StringComparer.OrdinalIgnoreCase))
                .ToList();
            if (familyRuns.Count == 0 || pdf.Options.HasNamedFontFamily(familyName)) continue;
            if (loadedFamilyCount >= MaximumLoadedSystemFontFamilies) break;
            if (!PdfCore.PdfEmbeddedFontFamily.TryFromSystem(familyName, out PdfCore.PdfEmbeddedFontFamily? family)
                || family == null) continue;

            if (!pdf.Options.TryRegisterNamedFontFamily(CreateCoverageStableFontFamily(family))) break;
            reservedFontSlots.Add(PdfCore.PdfStandardFontMapper.GetFontFamily(MapStandardFont(familyName)));
            loadedFamilyCount++;
        }
    }

    private static void RegisterLibrarySelectedDefaultSystemFontFamily(
        PdfCore.PdfDocument pdf,
        HtmlRenderDocument rendered,
        ISet<string> activeWebFontFamilies,
        ISet<PdfCore.PdfStandardFont> reservedFontSlots,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (pdf.Options.HasEmbeddedStandardFontFamily(PdfCore.PdfStandardFont.Helvetica)) {
            return;
        }

        List<HtmlRenderText> textRuns = EnumerateVisuals(
                rendered.Pages.SelectMany(page => page.Visuals))
            .OfType<HtmlRenderText>()
            .Where(text => !EnumerateFamilies(text.Font.FamilyName).Any(activeWebFontFamilies.Contains))
            .ToList();
        if (textRuns.Count == 0) {
            return;
        }

        foreach (string familyName in PdfCore.PdfOptions.DefaultDocumentFontFamilyFallback.Split(
                     new[] { ',', ';' },
                     StringSplitOptions.RemoveEmptyEntries)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!PdfCore.PdfEmbeddedFontFamily.TryFromSystem(
                    familyName.Trim(),
                    out PdfCore.PdfEmbeddedFontFamily? family)
                || family == null) {
                continue;
            }

            pdf.Options.RegisterFontFamily(
                PdfCore.PdfStandardFont.Helvetica,
                CreateCoverageSafeFontFamily(family, textRuns));
            reservedFontSlots.Add(PdfCore.PdfStandardFont.Helvetica);
            return;
        }
    }

    private static PdfCore.PdfEmbeddedFontFamily CreateCoverageStableFontFamily(
        PdfCore.PdfEmbeddedFontFamily family) {
        byte[] regular = family.Regular;
        byte[]? bold = SelectCoverageStableFace(family.Bold, regular);
        byte[]? italic = SelectCoverageStableFace(family.Italic, regular);
        byte[]? boldItalic = SelectCoverageStableFace(
            family.BoldItalic ?? family.Bold ?? family.Italic, regular);
        return new PdfCore.PdfEmbeddedFontFamily(family.FamilyName, regular, bold, italic, boldItalic);
    }

    private static byte[]? SelectCoverageStableFace(byte[]? styledFace, byte[] regularFace) {
        if (styledFace == null) return null;
        try {
            PdfCore.PdfTrueTypeFontProgram regular = PdfCore.PdfFontProgramCache.GetTrueType(regularFace, null);
            PdfCore.PdfTrueTypeFontProgram styled = PdfCore.PdfFontProgramCache.GetTrueType(styledFace, null);
            return regular.HasCoverageWithin(styled) ? styledFace : regularFace;
        } catch (Exception exception) when (PdfCore.PdfFontDiagnostics.IsFontProgramException(exception)) {
            return regularFace;
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

    private static IEnumerable<string> EnumerateUsedWebFontFamilyLists(
        IEnumerable<HtmlRenderVisual> visuals,
        OfficeFontFaceCollection faces) {
        foreach (HtmlRenderVisual visual in EnumerateVisuals(visuals)) {
            if (visual is HtmlRenderText text) {
                yield return text.Font.FamilyName;
            } else if (visual is HtmlRenderDrawing drawing) {
                foreach (string familyName in EnumerateDrawingWebFontFamilies(drawing.Drawing.Elements, faces)) {
                    yield return familyName;
                }
            }
        }
    }

    private static IEnumerable<string> EnumerateDrawingWebFontFamilies(
        IEnumerable<OfficeDrawingElement> elements,
        OfficeFontFaceCollection faces) {
        foreach (OfficeDrawingElement element in elements) {
            if (element is OfficeDrawingText text) {
                foreach (OfficeFontFallbackRun run in faces.PlanFallbackRuns(
                             text.Text,
                             text.Font.FamilyName,
                             text.Font.Style)) {
                    yield return run.FamilyName;
                }
            } else if (element is OfficeDrawingEffectGroup effectGroup) {
                foreach (string familyName in EnumerateDrawingWebFontFamilies(effectGroup.Drawing.Elements, faces)) {
                    yield return familyName;
                }
            } else if (element is OfficeDrawingTilingPattern tilingPattern) {
                foreach (string familyName in EnumerateDrawingWebFontFamilies(tilingPattern.Tile.Elements, faces)) {
                    yield return familyName;
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
                : visual is HtmlRenderLayoutRegion layoutRegion
                    ? layoutRegion.Visuals
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
        PdfCore.PdfTextDiagnostics.RequiresEmbeddedUnicodeFont(text);

    internal static string CollectRenderedTextForFontFallbackSelection(HtmlRenderDocument rendered) {
        var text = new System.Text.StringBuilder();
        foreach (HtmlRenderVisual visual in EnumerateVisuals(rendered.Pages.SelectMany(page => page.Visuals))) {
            if (visual is HtmlRenderText renderedText) text.Append(renderedText.Text);
            if (visual is HtmlRenderDrawing drawing) AppendDrawingText(drawing.Drawing.Elements, text);
        }
        return text.ToString();
    }

    private static void AppendDrawingText(IEnumerable<OfficeDrawingElement> elements, System.Text.StringBuilder text) {
        foreach (OfficeDrawingElement element in elements) {
            if (element is OfficeDrawingText drawingText) {
                text.Append(drawingText.Text);
            } else if (element is OfficeDrawingEffectGroup effectGroup) {
                AppendDrawingText(effectGroup.Drawing.Elements, text);
            } else if (element is OfficeDrawingTilingPattern tilingPattern) {
                AppendDrawingText(tilingPattern.Tile.Elements, text);
            }
        }
    }

    private static bool RegisterNamedFamily(
        PdfCore.PdfDocument pdf,
        string family,
        IReadOnlyList<OfficeFontFace> faces,
        CancellationToken cancellationToken) {
        OfficeFontFace regular = FindFace(faces, OfficeFontStyle.Regular) ?? faces[0];
        OfficeFontFace bold = FindFace(faces, OfficeFontStyle.Bold) ?? regular;
        OfficeFontFace italic = FindFace(faces, OfficeFontStyle.Italic) ?? regular;
        OfficeFontFace boldItalic = FindFace(faces, OfficeFontStyle.Bold | OfficeFontStyle.Italic) ?? bold;
        cancellationToken.ThrowIfCancellationRequested();
        return pdf.Options.TryRegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily(
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

    private static IEnumerable<string> EnumerateFamilies(string? familyNames) =>
        HtmlRenderCssValues.FontFamilyNames(familyNames);

    private static IEnumerable<string> EnumerateBoundedSystemFamilies(string? familyNames) =>
        HtmlRenderCssValues.TrySplitTopLevelCommas(
            familyNames, MaximumCssFontFamilyCandidates, out _)
            ? EnumerateFamilies(familyNames)
            : Enumerable.Empty<string>();

    private sealed class RegisteredWebFonts {
        internal RegisteredWebFonts(
            IReadOnlyDictionary<string, PdfCore.PdfStandardFont> slots,
            OfficeFontFaceCollection faces,
            PdfCore.PdfOptions options,
            HtmlDiagnosticReport diagnostics,
            OutlinedTextBudget outlineBudget,
            IOfficeTextShapingProvider? textShapingProvider,
            string? textShapingLanguage) {
            Slots = slots;
            Faces = faces;
            Options = options;
            Diagnostics = diagnostics;
            OutlineBudget = outlineBudget;
            TextShapingProvider = textShapingProvider;
            TextShapingLanguage = textShapingLanguage;
        }

        internal IReadOnlyDictionary<string, PdfCore.PdfStandardFont> Slots { get; }
        internal bool AllowInstalledFontFaces { get; set; }
        internal OfficeFontFaceCollection Faces { get; }
        internal PdfCore.PdfOptions Options { get; }
        internal HashSet<string> ReportedPrivateUseOmissions { get; } = new(StringComparer.Ordinal);
        internal Dictionary<string, bool> PrivateUsePaintability { get; } = new(StringComparer.Ordinal);
        internal HtmlDiagnosticReport Diagnostics { get; }
        internal OutlinedTextBudget OutlineBudget { get; }
        internal IOfficeTextShapingProvider? TextShapingProvider { get; }
        internal string? TextShapingLanguage { get; }
    }

    private sealed class OutlinedTextBudget {
        private int _remainingPathCommands;

        internal OutlinedTextBudget(
            int maximumCharactersPerRun,
            int maximumPathCommands) {
            if (maximumCharactersPerRun <= 0) throw new ArgumentOutOfRangeException(nameof(maximumCharactersPerRun));
            if (maximumPathCommands <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPathCommands));
            MaximumCharactersPerRun = maximumCharactersPerRun;
            _remainingPathCommands = maximumPathCommands;
            CffOperationBudget = new OfficeCffOperationBudget();
        }

        internal int MaximumCharactersPerRun { get; }

        internal OfficeCffOperationBudget CffOperationBudget { get; }

        internal int RemainingPointAllowance {
            get {
                if (_remainingPathCommands <= 0) {
                    throw new InvalidOperationException("HTML-to-PDF outlined text exceeded the configured path-command budget.");
                }
                return _remainingPathCommands;
            }
        }

        internal void ValidateTextLength(int characterCount) {
            if (characterCount > MaximumCharactersPerRun) {
                throw new InvalidOperationException("HTML-to-PDF outlined text exceeded the configured per-run character budget.");
            }
        }

        internal void ConsumePathCommand() {
            if (_remainingPathCommands <= 0) {
                throw new InvalidOperationException("HTML-to-PDF outlined text exceeded the configured path-command budget.");
            }
            _remainingPathCommands--;
        }
    }
}
