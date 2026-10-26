using System.Collections.Generic;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private sealed class NativeFontMap {
            private readonly Dictionary<string, PdfCore.PdfStandardFont> _fontSlots = new(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, string> _namedFontFamilies = new(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _reportedFontSubstitution = new(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, IReadOnlyList<string>> _declaredFontAlternates = new(StringComparer.OrdinalIgnoreCase);
            private readonly PdfCore.PdfConversionReport? _report;
            private PdfCore.PdfOptions? _resolvedOptions;

            public NativeFontMap() : this(null) { }

            public NativeFontMap(PdfCore.PdfConversionReport? report, WordDocument? document = null) {
                _report = report;
                if (document?._wordprocessingDocument?.MainDocumentPart?.FontTablePart?.Fonts is W.Fonts fonts) {
                    foreach (W.Font font in fonts.Elements<W.Font>()) {
                        string? primary = font.Name?.Value;
                        string? alternateNames = font.GetFirstChild<W.AltName>()?.Val?.Value;
                        if (string.IsNullOrWhiteSpace(primary) || string.IsNullOrWhiteSpace(alternateNames)) continue;
                        string key = NormalizeNativeFontFamily(primary!);
                        if (_declaredFontAlternates.ContainsKey(key)) continue;
                        int maximumAlternates = OfficeFontFamilyParser.DefaultMaximumCandidates - 1;
                        var alternates = new List<string>(maximumAlternates);
                        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { key };
                        // Word treats the complete value as one alternate name; retain
                        // parsed fallback ordering only after that literal lookup.
                        string literalAlternate = alternateNames!.Trim();
                        if (literalAlternate.Length <= OfficeFontFamilyParser.DefaultMaximumFamilyNameLength &&
                            seen.Add(NormalizeNativeFontFamily(literalAlternate))) {
                            alternates.Add(literalAlternate);
                        }
                        foreach (string alternate in OfficeFontFamilyParser.Parse(alternateNames,
                                     maximumCandidates: maximumAlternates)) {
                            if (alternates.Count >= maximumAlternates) break;
                            if (seen.Add(NormalizeNativeFontFamily(alternate))) alternates.Add(alternate);
                        }
                        if (alternates.Count > 0) _declaredFontAlternates[key] = alternates;
                    }
                }
            }

            public IReadOnlyList<string> GetDeclaredFontAlternates(string? familyName) =>
                !string.IsNullOrWhiteSpace(familyName) &&
                _declaredFontAlternates.TryGetValue(NormalizeNativeFontFamily(familyName!), out IReadOnlyList<string>? alternates)
                    ? alternates : Array.Empty<string>();

            public bool UsePdfDefaultForDocumentDefaultFont { get; private set; }

            public double? DefaultFontSize => _resolvedOptions?.DefaultFontSize;

            public double? MeasureText(PdfCore.PdfTextRun run) => _resolvedOptions == null
                ? null : PdfCore.PdfWriter.MeasurePositionedText(run, _resolvedOptions);

            public PdfCore.PdfWriter.PositionedTextBoxMeasurement? MeasureTextBox(PdfCore.PdfTextRun run) => _resolvedOptions == null
                ? null : PdfCore.PdfWriter.MeasurePositionedTextBox(run, _resolvedOptions);

            public PdfCore.PdfWriter.TextDecorationMeasurement? MeasureStrikethrough(PdfCore.PdfTextRun run,
                double pointsPerPixel) => _resolvedOptions == null ? null :
                    PdfCore.PdfWriter.MeasureStrikethrough(run, _resolvedOptions, pointsPerPixel);

            public PdfCore.PdfWriter.TextDecorationMeasurement? MeasureUnderline(PdfCore.PdfTextRun run,
                double pointsPerPixel) => _resolvedOptions == null ? null :
                    PdfCore.PdfWriter.MeasureUnderline(run, _resolvedOptions, pointsPerPixel);

            public void ReportCheckBoxAppearanceFallback(string fieldName, bool isChecked) =>
                _report?.Add(new PdfCore.PdfConversionWarning("OfficeIMO.Word.Pdf",
                    "NativeCheckBoxGlyphAppearanceUnavailable", "word:contentControl[" + fieldName + "]",
                    "The authored " + (isChecked ? "checked" : "unchecked") +
                    " glyph could not be represented by the available PDF fonts; the field uses the standard check box appearance for that state."));

            public void ReportInlineControlStaticBackground(string fieldName) =>
                _report?.Add(new PdfCore.PdfConversionWarning("OfficeIMO.Word.Pdf",
                    "NativeInlineControlBackgroundStatic", "word:contentControl[" + fieldName + "]",
                    "The content control retains static text so its run background stays beneath surrounding paragraph text; it is not an editable PDF field.",
                    severity: PdfCore.PdfConversionWarningSeverity.Warning,
                    lossKind: OfficeIMO.OfficeConversionLossKind.Omission));

            public PdfCore.PdfWriter.CanvasTextBoxMeasurement? MeasureCanvasTextBox(
                IReadOnlyList<PdfCore.PdfTextRun> runs, double width, double height, PdfCore.PdfCanvasTextBoxStyle style) =>
                _resolvedOptions == null ? null : PdfCore.PdfWriter.MeasureCanvasTextBox(runs, width, height, style, _resolvedOptions);

            public void PreferPdfDefaultForDocumentDefaultFont() =>
                UsePdfDefaultForDocumentDefaultFont = true;

            public void AttachResolvedOptions(PdfCore.PdfOptions options) =>
                _resolvedOptions = options ?? throw new ArgumentNullException(nameof(options));

            public void Register(string familyName, PdfCore.PdfStandardFont fontSlot) {
                if (string.IsNullOrWhiteSpace(familyName)) {
                    return;
                }

                _fontSlots[NormalizeNativeFontFamily(familyName)] = PdfCore.PdfStandardFontMapper.GetFontFamily(fontSlot);
            }

            public bool TryGetFontSlot(string? familyName, out PdfCore.PdfStandardFont fontSlot) {
                fontSlot = PdfCore.PdfStandardFont.Helvetica;
                return !string.IsNullOrWhiteSpace(familyName) &&
                    _fontSlots.TryGetValue(NormalizeNativeFontFamily(familyName!), out fontSlot);
            }

            public void RegisterNamed(string familyName, string registeredFamilyName) {
                if (string.IsNullOrWhiteSpace(familyName) || string.IsNullOrWhiteSpace(registeredFamilyName)) {
                    return;
                }

                _namedFontFamilies[NormalizeNativeFontFamily(familyName)] = registeredFamilyName.Trim();
            }

            public bool TryGetNamedFontFamily(string? familyName, out string? registeredFamilyName) {
                registeredFamilyName = null;
                if (string.IsNullOrWhiteSpace(familyName) ||
                    !_namedFontFamilies.TryGetValue(NormalizeNativeFontFamily(familyName!), out string? rawFamilyName)) return false;
                // PDF font-family properties accept lists; a resolved document
                // family remains one literal name even when it contains a comma.
                registeredFamilyName = OfficeFontFamilyParser.Quote(rawFamilyName);
                return true;
            }

            public bool TryResolveLineSpacingRatio(string? familyName, out double ratio) {
                ratio = 0D;
                if (_resolvedOptions == null || string.IsNullOrWhiteSpace(familyName)) {
                    return false;
                }

                if (_namedFontFamilies.TryGetValue(NormalizeNativeFontFamily(familyName!), out string? registeredFamilyName) &&
                    TryResolveNamedLineSpacingRatio(_resolvedOptions, registeredFamilyName, out ratio)) {
                    return true;
                }

                if (TryGetFontSlot(familyName, out PdfCore.PdfStandardFont fontSlot) ||
                    PdfCore.PdfStandardFontMapper.TryMapFontFamily(OfficeFontFamilyParser.Quote(familyName!), out fontSlot)) {
                    return TryResolveSlotLineSpacingRatio(_resolvedOptions, fontSlot, out ratio);
                }

                return false;
            }

            public bool TryResolveDefaultLineSpacingRatio(out double ratio) {
                ratio = 0D;
                return _resolvedOptions != null &&
                       TryResolveSlotLineSpacingRatio(_resolvedOptions, _resolvedOptions.DefaultFont, out ratio);
            }

            private static bool TryResolveNamedLineSpacingRatio(
                PdfCore.PdfOptions options,
                string? registeredFamilyName,
                out double ratio) {
                ratio = 0D;
                if (string.IsNullOrWhiteSpace(registeredFamilyName)) {
                    return false;
                }

                if (options.NamedFontFamilies.TryGetValue(
                        registeredFamilyName!,
                        out PdfCore.PdfEmbeddedFontFamily? family)) {
                    return TryResolveEmbeddedLineSpacingRatio(
                        family.RegularSnapshot,
                        out ratio);
                }

                if (!options.TryGetRenderingProfileFamilyFallbacks(
                        OfficeFontFamilyParser.Quote(registeredFamilyName!),
                        out PdfCore.PdfEmbeddedFontFallbackSet? planner)
                    || planner == null) {
                    return false;
                }

                bool resolved = false;
                double maximumRatio = 0D;
                foreach (PdfCore.PdfEmbeddedFontFallbackCandidate candidate in
                         planner.Candidates) {
                    if (TryResolveEmbeddedLineSpacingRatio(
                            candidate.DataSnapshot,
                            out double candidateRatio)) {
                        maximumRatio = Math.Max(maximumRatio, candidateRatio);
                        resolved = true;
                    }
                }

                ratio = maximumRatio;
                return resolved;
            }

            private static bool TryResolveSlotLineSpacingRatio(
                PdfCore.PdfOptions options,
                PdfCore.PdfStandardFont slot,
                out double ratio) {
                PdfCore.PdfStandardFont normalized = PdfCore.PdfStandardFontMapper.GetFontFamily(slot);
                if (options.EmbeddedFonts.TryGetValue(normalized, out PdfCore.PdfEmbeddedFont? embedded) &&
                    TryResolveEmbeddedLineSpacingRatio(embedded.DataSnapshot, out ratio)) {
                    return true;
                }

                // Standard PDF fonts expose no portable font program to inspect.
                ratio = NativeDefaultParagraphLineHeight;
                return true;
            }

            private static bool TryResolveEmbeddedLineSpacingRatio(byte[] fontData, out double ratio) {
                OfficeOpenTypeLineMetrics? metrics = OfficeOpenTypeLineMetrics.TryRead(fontData);
                ratio = metrics?.HorizontalAdvanceRatio ?? 0D;
                return ratio >= 0.8D && ratio <= 2D;
            }

            public void ReportSlotExhaustion(string familyName, PdfCore.PdfStandardFont fallbackSlot, string? occupyingFontFamily) {
                string normalizedFamily = NormalizeNativeFontFamily(familyName);
                if (_report == null || !_reportedFontSubstitution.Add(normalizedFamily)) {
                    return;
                }

                PdfCore.PdfStandardFont normalizedSlot = PdfCore.PdfStandardFontMapper.GetFontFamily(fallbackSlot);
                string message = string.IsNullOrWhiteSpace(occupyingFontFamily)
                    ? "The installed font family could not receive a distinct embedded PDF family slot because all standard-family slots are occupied; runs use the mapped PDF family " + normalizedSlot + "."
                    : "The installed font family could not receive a distinct embedded PDF family slot because all standard-family slots are occupied; runs use the occupying embedded family '" + occupyingFontFamily + "' in the logical " + normalizedSlot + " slot.";
                var details = new Dictionary<string, string> {
                    ["fontFamily"] = familyName,
                    ["fallbackSlot"] = normalizedSlot.ToString()
                };
                if (!string.IsNullOrWhiteSpace(occupyingFontFamily)) {
                    details["occupyingFontFamily"] = occupyingFontFamily!;
                }

                _report.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.Word.Pdf",
                    "NativeFontFamilySlotExhausted",
                    "word:font[" + familyName + "]",
                    message,
                    details: details));
            }

            public void ReportFontSubstitution(
                PdfCore.PdfOptions options,
                string familyName,
                PdfCore.PdfStandardFont? fallbackSlot,
                string? resolvedFontFamily = null,
                bool sourceFontEmbeddingAllowed = true) {
                string normalizedFamily = NormalizeNativeFontFamily(familyName);
                if (_report == null || !_reportedFontSubstitution.Add(normalizedFamily)) {
                    return;
                }

                _report.Add(options.CreateFontFamilySubstitutionWarning(
                    "OfficeIMO.Word.Pdf",
                    "NativeFontFamilySubstituted",
                    "word:font[" + familyName + "]",
                    familyName,
                    fallbackSlot,
                    resolvedFontFamily,
                    sourceFontEmbeddingAllowed: sourceFontEmbeddingAllowed));
            }
        }

        private static W.RunFonts? GetNativeEmptyParagraphMarkFonts(WordParagraph paragraph, IReadOnlyList<WordParagraph> runs) =>
            runs.Any(run => IsNativeParagraphMetricBearingRun(run, paragraph, paragraph._paragraph.Ancestors<W.FootnoteEndnoteType>().Any()))
                ? null : paragraph._paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<W.RunFonts>();

        private static void RegisterNativeThemeStyleFonts(
            WordDocument document,
            PdfCore.PdfOptions pdfOptions,
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
            bool allowSystemFontEmbedding,
            NativeFontMap nativeFontMap) {
            var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (WordSection section in document.Sections) {
                foreach (WordElement element in EnumerateNativeTableOfContentsElements(section)) {
                    if (element is not WordParagraph paragraph ||
                        GetNativeTableOfContentsHeadingLevel(paragraph) <= 0) {
                        continue;
                    }

                    RegisterNativeFontCandidates(
                        ResolveNativeParagraphStyleFontFamilies(document, paragraph.StyleId),
                        pdfOptions,
                        registeredFamilies,
                        registeredFontSlots,
                        allowSystemFontEmbedding,
                        nativeFontMap);
                }
            }
        }

        private static IEnumerable<string> ResolveNativeParagraphStyleFontFamilies(WordDocument? document, string? styleId) {
            IReadOnlyList<W.Style> styleChain = GetNativeParagraphStyleChain(document, styleId);
            NativeLatinFontFamilies families = default;
            foreach (W.Style style in styleChain) {
                W.RunFonts? runFonts = style.GetFirstChild<W.StyleRunProperties>()?.GetFirstChild<W.RunFonts>();
                families = GetNativeRunFontFamilies(document, runFonts).Inherit(families);
            }

            return families.Enumerate();
        }

        private static string? ResolveNativeRunFontsFamily(WordDocument? document, W.RunFonts? runFonts) =>
            EnumerateNativeLatinFontFamilies(document, runFonts).FirstOrDefault();

        private static IEnumerable<string> EnumerateNativeLatinFontFamilies(WordDocument? document, W.RunFonts? runFonts) =>
            GetNativeRunFontFamilies(document, runFonts).Enumerate();

        private static string? GetNativeThemeFontValue(DocumentFormat.OpenXml.EnumValue<W.ThemeFontValues>? value) {
            if (value == null) {
                return null;
            }

            return string.IsNullOrWhiteSpace(value.InnerText) ? value.Value.ToString() : value.InnerText;
        }

        private static string? ResolveNativeThemeFontFamily(WordDocument? document, params string?[] themeValues) {
            A.FontScheme? fontScheme = document?._wordprocessingDocument?.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;
            if (fontScheme == null) {
                return null;
            }

            foreach (string? themeValue in themeValues) {
                if (string.IsNullOrWhiteSpace(themeValue)) {
                    continue;
                }

                string normalized = NormalizeNativeFontFamily(themeValue!);
                if (normalized.StartsWith("major", StringComparison.OrdinalIgnoreCase)) {
                    return fontScheme.MajorFont?.LatinFont?.Typeface?.Value;
                }

                if (normalized.StartsWith("minor", StringComparison.OrdinalIgnoreCase)) {
                    return fontScheme.MinorFont?.LatinFont?.Typeface?.Value;
                }
            }

            return null;
        }

        private static string? FirstNonWhiteSpace(params string?[] values) {
            foreach (string? value in values) {
                if (!string.IsNullOrWhiteSpace(value)) {
                    return value;
                }
            }

            return null;
        }

        private static string NormalizeNativeFontFamily(string familyName) {
            return PdfCore.PdfOptions.NormalizeOfficeFontFamilyKey(familyName);
        }
    }
}
