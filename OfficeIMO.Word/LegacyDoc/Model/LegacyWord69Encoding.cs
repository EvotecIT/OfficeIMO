using OfficeIMO.Core.Internal;

namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Decodes qualified single-byte text using the active legacy font and language.</summary>
    internal static class LegacyWord69Encoding {
        internal static LegacyDocTextCharacter[] DecodeText(LegacyWord69Bytes text, int textStart,
            IReadOnlyList<byte> charsets, IReadOnlyList<TextContextRange> contexts,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphs, LegacyDocStyleSheet stylesheet,
            IReadOnlyDictionary<ushort, int> styleFonts, IReadOnlyDictionary<ushort, ushort> styleLanguages,
            int defaultFont, ushort defaultLanguage, out bool unsupported,
            System.Threading.CancellationToken cancellationToken = default) {
            unsupported = false;
            var characters = new LegacyDocTextCharacter[text.Length];
            bool inspectContexts = charsets.Count != 0 || contexts.Count != 0 || styleFonts.Count != 0 || styleLanguages.Count != 0;

            var resolvedFonts = new Dictionary<ushort, int>();
            var resolvedLanguages = new Dictionary<ushort, ushort>();
            bool mayUseSymbol = charsets.Contains((byte)2);
            int characterIndex = 0;
            int paragraphIndex = 0;
            for (int i = 0; i < text.Length; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                byte value = text[i];
                int offset = textStart + i;
                char decoded = LegacyDocPieceTable.DecodeWindows1252(value);
                if (!inspectContexts || value < 0x20 || value == 0x7F || (!mayUseSymbol && value < 0x80)) {
                    characters[i] = new LegacyDocTextCharacter(decoded, offset, i);
                    continue;
                }

                while (characterIndex < contexts.Count && contexts[characterIndex].End <= offset) characterIndex++;
                while (paragraphIndex < paragraphs.Count && paragraphs[paragraphIndex].FileOffsetEnd <= offset) paragraphIndex++;
                TextContextRange context = characterIndex < contexts.Count && contexts[characterIndex].Start <= offset
                    ? contexts[characterIndex]
                    : default;
                ushort style = paragraphIndex < paragraphs.Count && paragraphs[paragraphIndex].FileOffsetStart <= offset
                    ? paragraphs[paragraphIndex].Format.StyleIndex ?? 0
                    : (ushort)0;

                int font = context.FontIndex ?? ResolveStyleValue(style, stylesheet, styleFonts, defaultFont, resolvedFonts);
                byte charset = font >= 0 && font < charsets.Count ? charsets[font] : (byte)0;
                // Symbol assigns non-Latin meaning to ASCII positions. Word 6/95 also uses the
                // text language when the selected font has an unresolved charset (0 or 1).
                if (charset == 2) {
                    unsupported = true;
                } else if (value >= 0x80) {
                    if (charset == 238) {
                        // EASTEUROPE_CHARSET explicitly selects Windows-1250, even when the
                        // document's default language is Western or the font is inherited.
                        decoded = OfficeLegacySingleByteEncoding.DecodeByte(value, 1250);
                    } else if (charset == 186) {
                        // BALTIC_CHARSET explicitly selects Windows-1257 for this text run.
                        decoded = OfficeLegacySingleByteEncoding.DecodeByte(value, 1257);
                    } else if (charset > 1 || (context.FontIndex.HasValue && font >= charsets.Count)) {
                        unsupported = true;
                    } else {
                        ushort language = context.LanguageId ?? ResolveStyleValue(style, stylesheet, styleLanguages, defaultLanguage, resolvedLanguages);
                        if (language == 0x0405) {
                            // The qualified OpenOffice Word 6 source uses Czech text with
                            // both EASTEUROPE_CHARSET and unresolved ANSI font charsets.
                            decoded = OfficeLegacySingleByteEncoding.DecodeByte(value, 1250);
                        } else if (!LegacyWord69Reader.IsWesternLanguage(language)) {
                            unsupported = true;
                        }
                    }
                }
                characters[i] = new LegacyDocTextCharacter(decoded, offset, i);
            }

            return characters;
        }

        private static T ResolveStyleValue<T>(ushort style, LegacyDocStyleSheet stylesheet,
            IReadOnlyDictionary<ushort, T> values, T fallback, Dictionary<ushort, T> cache) where T : struct {
            if (cache.TryGetValue(style, out T cached)) return cached;

            T value = fallback;
            var chain = new HashSet<ushort>();
            while (chain.Add(style)) {
                if (cache.TryGetValue(style, out T resolved) || values.TryGetValue(style, out resolved)) {
                    value = resolved;
                    break;
                }

                if (stylesheet.TryGetParagraphStyle(style, out var definition) && definition.BasedOnStyleIndex is ushort baseStyle) {
                    style = baseStyle;
                } else if (style != 0) {
                    style = 0;
                } else {
                    break;
                }
            }

            foreach (ushort index in chain) cache[index] = value;
            return value;
        }

        internal readonly record struct TextContextRange(int Start, int End, int? FontIndex, ushort? LanguageId);
    }
}
