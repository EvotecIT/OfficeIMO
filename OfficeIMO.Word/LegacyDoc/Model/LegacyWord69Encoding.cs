namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Checks the active legacy font and language before interpreting text bytes as Windows-1252.</summary>
    internal static class LegacyWord69Encoding {
        internal static bool HasUnsupportedText(LegacyWord69Bytes text, int textStart,
            IReadOnlyList<byte> charsets, IReadOnlyList<TextContextRange> contexts,
            IReadOnlyList<LegacyDocParagraphFormatRange> paragraphs, LegacyDocStyleSheet stylesheet,
            IReadOnlyDictionary<ushort, int> styleFonts, IReadOnlyDictionary<ushort, ushort> styleLanguages,
            int defaultFont, ushort defaultLanguage) {
            if (charsets.Count == 0 && contexts.Count == 0 && styleFonts.Count == 0 && styleLanguages.Count == 0) return false;

            var resolvedFonts = new Dictionary<ushort, int>();
            var resolvedLanguages = new Dictionary<ushort, ushort>();
            bool mayUseSymbol = charsets.Contains((byte)2);
            int characterIndex = 0;
            int paragraphIndex = 0;
            for (int i = 0; i < text.Length; i++) {
                byte value = text[i];
                if (value < 0x20 || value == 0x7F || (!mayUseSymbol && value < 0x80)) continue;

                int offset = textStart + i;
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
                if (charset == 2) return true;
                if (value < 0x80) continue;
                if (charset > 1 || (context.FontIndex.HasValue && font >= charsets.Count)) return true;
                ushort language = context.LanguageId ?? ResolveStyleValue(style, stylesheet, styleLanguages, defaultLanguage, resolvedLanguages);
                if (!LegacyWord69Reader.IsWesternLanguage(language)) return true;
            }

            return false;
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