using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private sealed class NativeFontMap {
            private readonly Dictionary<string, PdfCore.PdfStandardFont> _fontSlots = new(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _reportedSlotExhaustion = new(StringComparer.OrdinalIgnoreCase);
            private readonly PdfCore.PdfConversionReport? _report;

            public NativeFontMap() : this(null) { }

            public NativeFontMap(PdfCore.PdfConversionReport? report) {
                _report = report;
            }

            public bool UsePdfDefaultForDocumentDefaultFont { get; private set; }

            public void PreferPdfDefaultForDocumentDefaultFont() =>
                UsePdfDefaultForDocumentDefaultFont = true;

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

            public void ReportSlotExhaustion(string familyName, PdfCore.PdfStandardFont fallbackSlot, string occupyingFontFamily) {
                string normalizedFamily = NormalizeNativeFontFamily(familyName);
                if (_report == null || !_reportedSlotExhaustion.Add(normalizedFamily)) {
                    return;
                }

                PdfCore.PdfStandardFont normalizedSlot = PdfCore.PdfStandardFontMapper.GetFontFamily(fallbackSlot);
                _report.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.Word.Pdf",
                    "NativeFontFamilySlotExhausted",
                    "word:font[" + familyName + "]",
                    "The installed font family could not receive a distinct embedded PDF family slot because all standard-family slots are occupied; runs use the occupying embedded family '" + occupyingFontFamily + "' in the logical " + normalizedSlot + " slot.",
                    details: new Dictionary<string, string> {
                        ["fontFamily"] = familyName,
                        ["fallbackSlot"] = normalizedSlot.ToString(),
                        ["occupyingFontFamily"] = occupyingFontFamily
                    }));
            }
        }

        private static void RegisterNativeThemeStyleFonts(
            WordDocument document,
            PdfCore.PdfOptions pdfOptions,
            HashSet<PdfCore.PdfStandardFont> registeredFontSlots,
            bool allowSystemFontEmbedding,
            NativeFontMap nativeFontMap) {
            if (!allowSystemFontEmbedding) {
                return;
            }

            var registeredFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string styleId in new[] { "Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6", "Heading7", "Heading8", "Heading9" }) {
                string? familyName = ResolveNativeParagraphStyleFontFamily(document, styleId);
                if (!PdfCore.PdfOptions.TryAddOfficeFontFamilyKey(familyName, registeredFamilies, NormalizeNativeFontFamily, out string trimmedFamilyName)) {
                    continue;
                }

                PdfCore.PdfStandardFont slot = SelectNativeAdditionalFontSlot(trimmedFamilyName, pdfOptions, registeredFontSlots);
                bool slotAlreadyEmbedded = pdfOptions.HasEmbeddedStandardFontFamily(slot);
                if (!slotAlreadyEmbedded) {
                    pdfOptions.RegisterOfficeFontFamily(trimmedFamilyName, slot, embedSystemFont: true);
                }

                if (slotAlreadyEmbedded || pdfOptions.HasEmbeddedStandardFontFamily(slot)) {
                    registeredFontSlots.Add(slot);
                    nativeFontMap.Register(trimmedFamilyName, slot);
                    continue;
                }

                if (PdfCore.PdfStandardFontMapper.TryMapFontFamily(trimmedFamilyName, out PdfCore.PdfStandardFont mappedFont)) {
                    PdfCore.PdfStandardFont mappedFamily = PdfCore.PdfStandardFontMapper.GetFontFamily(mappedFont);
                    registeredFontSlots.Add(mappedFamily);
                    nativeFontMap.Register(trimmedFamilyName, mappedFamily);
                }
            }
        }

        private static PdfCore.PdfStandardFont SelectNativeAdditionalFontSlot(string familyName, PdfCore.PdfOptions pdfOptions, HashSet<PdfCore.PdfStandardFont> registeredFontSlots) {
            if (PdfCore.PdfOptions.TrySelectAvailableFontFamilySlot(familyName, registeredFontSlots, out PdfCore.PdfStandardFont fontSlot)) {
                return fontSlot;
            }

            return PdfCore.PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfCore.PdfStandardFont mappedFont)
                ? PdfCore.PdfStandardFontMapper.GetFontFamily(mappedFont)
                : PdfCore.PdfStandardFont.Helvetica;
        }

        private static string? ResolveNativeParagraphStyleFontFamily(WordDocument? document, string? styleId) {
            IReadOnlyList<W.Style> styleChain = GetNativeParagraphStyleChain(document, styleId);
            string? familyName = null;
            foreach (W.Style style in styleChain) {
                W.RunFonts? runFonts = style.GetFirstChild<W.StyleRunProperties>()?.GetFirstChild<W.RunFonts>();
                familyName = ResolveNativeRunFontsFamily(document, runFonts) ?? familyName;
            }

            return familyName;
        }

        private static string? ResolveNativeRunFontsFamily(WordDocument? document, W.RunFonts? runFonts) {
            if (runFonts == null) {
                return null;
            }

            string? directFamily = FirstNonWhiteSpace(runFonts.Ascii?.Value, runFonts.HighAnsi?.Value);
            if (!string.IsNullOrWhiteSpace(directFamily)) {
                return directFamily;
            }

            string? themeFamily = ResolveNativeThemeFontFamily(
                document,
                GetNativeThemeFontValue(runFonts.AsciiTheme),
                GetNativeThemeFontValue(runFonts.HighAnsiTheme));
            return string.IsNullOrWhiteSpace(themeFamily) ? null : themeFamily;
        }

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
            return familyName.Trim().Replace(" ", string.Empty).Replace("-", string.Empty);
        }
    }
}
