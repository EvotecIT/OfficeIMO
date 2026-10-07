using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        // A declaration overrides one logical slot; the other slot still inherits.
        // Availability is decided by the existing PDF font/resource policy.
        private readonly record struct NativeLatinFontFamilies(string? Ascii, string? HighAnsi) {
            public string? Primary => Ascii ?? HighAnsi;
            public NativeLatinFontFamilies Inherit(NativeLatinFontFamilies inherited) =>
                new(Ascii ?? inherited.Ascii, HighAnsi ?? inherited.HighAnsi);

            public IEnumerable<string> Enumerate() {
                if (Ascii != null) yield return Ascii;
                if (HighAnsi != null) yield return HighAnsi;
            }
        }

        private static NativeLatinFontFamilies GetNativeRunFontFamilies(WordDocument? document, W.RunFonts? fonts) =>
            fonts == null ? default : new(
                FirstNonWhiteSpace(ResolveNativeThemeFontFamily(document, GetNativeThemeFontValue(fonts.AsciiTheme)), fonts.Ascii?.Value),
                FirstNonWhiteSpace(ResolveNativeThemeFontFamily(document, GetNativeThemeFontValue(fonts.HighAnsiTheme)), fonts.HighAnsi?.Value));

        private static IEnumerable<string> EnumerateNativeFontFamilies(string? primary, NativeLatinFontFamilies families) =>
            families.Primary != null ? families.Enumerate() :
                (string.IsNullOrWhiteSpace(primary) ? Array.Empty<string>() : new[] { primary! });

        private static IEnumerable<string> EnumerateNativeStyleFontFamilies(
            NativeCharacterStyleDefaults character, NativeParagraphStyleDefaults paragraph,
            NativeTableRunStyleDefaults table, NativeDocumentDefaults document, bool includeDocument = true) =>
            EnumerateNativeFontFamilies(character.FontFamily, character.FontFamilies)
                .Concat(EnumerateNativeFontFamilies(paragraph.FontFamily, paragraph.FontFamilies))
                .Concat(EnumerateNativeFontFamilies(table.FontFamily, table.FontFamilies))
                .Concat(includeDocument ? EnumerateNativeFontFamilies(document.FontFamily, document.FontFamilies) : Array.Empty<string>());

        private static void RegisterNativeFontCandidates(
            IEnumerable<string> families, PdfCore.PdfOptions options, HashSet<string> registeredFamilies,
            HashSet<PdfCore.PdfStandardFont> registeredSlots, bool allowSystemFontEmbedding, NativeFontMap map) {
            foreach (string family in families) {
                RegisterNativeFontCandidate(family, options, registeredFamilies, registeredSlots, allowSystemFontEmbedding, map);
                if (map.TryGetFontSlot(family, out _) || map.TryGetNamedFontFamily(family, out _)) return;
            }
        }
    }
}
