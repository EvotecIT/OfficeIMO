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
            private readonly PdfCore.PdfConversionReport? _report;
            private PdfCore.PdfOptions? _resolvedOptions;

            public NativeFontMap() : this(null) { }

            public NativeFontMap(PdfCore.PdfConversionReport? report) {
                _report = report;
            }

            public bool UsePdfDefaultForDocumentDefaultFont { get; private set; }

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
                return !string.IsNullOrWhiteSpace(familyName) &&
                       _namedFontFamilies.TryGetValue(NormalizeNativeFontFamily(familyName!), out registeredFamilyName);
            }

            public bool TryResolveLineSpacingRatio(string? familyName, out double ratio) {
                ratio = 0D;
                if (_resolvedOptions == null || string.IsNullOrWhiteSpace(familyName)) {
                    return false;
                }

                if (TryGetNamedFontFamily(familyName, out string? registeredFamilyName) &&
                    TryResolveNamedLineSpacingRatio(_resolvedOptions, registeredFamilyName, out ratio)) {
                    return true;
                }

                if (TryGetFontSlot(familyName, out PdfCore.PdfStandardFont fontSlot) ||
                    PdfCore.PdfStandardFontMapper.TryMapFontFamily(familyName, out fontSlot)) {
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
                if (string.IsNullOrWhiteSpace(registeredFamilyName) ||
                    !options.NamedFontFamilies.TryGetValue(registeredFamilyName!, out PdfCore.PdfEmbeddedFontFamily? family)) {
                    return false;
                }

                return TryResolveEmbeddedLineSpacingRatio(family.RegularSnapshot, out ratio);
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
                OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoad(fontData);
                ratio = font?.LineSpacingRatio ?? 0D;
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

            public void ReportFontSubstitution(string familyName, PdfCore.PdfStandardFont fallbackSlot, string? resolvedFontFamily = null) {
                string normalizedFamily = NormalizeNativeFontFamily(familyName);
                if (_report == null || !_reportedFontSubstitution.Add(normalizedFamily)) {
                    return;
                }

                PdfCore.PdfStandardFont normalizedSlot = PdfCore.PdfStandardFontMapper.GetFontFamily(fallbackSlot);
                string message = string.IsNullOrWhiteSpace(resolvedFontFamily)
                    ? "The source font family '" + familyName + "' was unavailable or could not be embedded; generated text uses the mapped PDF family " + normalizedSlot + "."
                    : "The source font family '" + familyName + "' was unavailable or could not be embedded; generated text uses the embedded family '" + resolvedFontFamily + "' through the logical " + normalizedSlot + " PDF slot.";
                var details = new Dictionary<string, string> {
                    ["fontFamily"] = familyName,
                    ["fallbackSlot"] = normalizedSlot.ToString()
                };
                if (!string.IsNullOrWhiteSpace(resolvedFontFamily)) {
                    details["resolvedFontFamily"] = resolvedFontFamily!;
                }

                _report.Add(new PdfCore.PdfConversionWarning(
                    "OfficeIMO.Word.Pdf",
                    "NativeFontFamilySubstituted",
                    "word:font[" + familyName + "]",
                    message,
                    details: details));
            }
        }

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

                    RegisterNativeFontCandidate(
                        ResolveNativeParagraphStyleFontFamily(document, paragraph.StyleId),
                        pdfOptions,
                        registeredFamilies,
                        registeredFontSlots,
                        allowSystemFontEmbedding,
                        nativeFontMap);
                }
            }
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
            return PdfCore.PdfOptions.NormalizeOfficeFontFamilyKey(familyName);
        }
    }
}
