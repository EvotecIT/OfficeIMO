using System.Runtime.CompilerServices;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const int MaxNativeIndexedStyles = 4_096;
        private const int MaxNativeStyleChainDepth = 256;
        private const int MaxNativeCachedStyleReferences = 8_192;
        private static readonly ConditionalWeakTable<WordDocument, NativeStyleLookupCache> NativeStyleLookupCaches = new();

        private sealed class NativeStyleLookupCache {
            internal NativeStyleIndex ParagraphStyles { get; } = new();
            internal NativeStyleIndex CharacterStyles { get; } = new();
            internal NativeStyleIndex TableStyles { get; } = new();
            internal Dictionary<string, IReadOnlyList<W.Style>> ParagraphChains { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, IReadOnlyList<W.Style>> CharacterChains { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, IReadOnlyList<W.Style>> TableChains { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, NativeParagraphStyleDefaults> ParagraphDefaults { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, NativeCharacterStyleDefaults> CharacterDefaults { get; } = new(StringComparer.Ordinal);
            internal string? DefaultParagraphStyleId { get; }
            internal string? DefaultTableStyleId { get; }
            private int _indexedStyleCount;
            private int _cachedStyleReferenceCount;

            internal NativeStyleLookupCache(WordStyleCatalog styles) {
                string? defaultParagraphStyleId = null;
                string? defaultTableStyleId = null;
                foreach (WordStyleDefinitionsReader.StyleDefinition style in styles.Definitions) {
                    string? styleId = style.StyleId;
                    if (string.IsNullOrEmpty(styleId)) {
                        continue;
                    }

                    if (IsNativeStyleType(style, "paragraph")) {
                        if (!ParagraphStyles.ContainsKey(styleId!)) {
                            RecordIndexedStyle();
                            ParagraphStyles.Add(styleId!, style);
                            if (defaultParagraphStyleId == null && style.IsDefault) {
                                defaultParagraphStyleId = styleId;
                            }
                        }
                    } else if (IsNativeStyleType(style, "character") && !CharacterStyles.ContainsKey(styleId!)) {
                        RecordIndexedStyle();
                        CharacterStyles.Add(styleId!, style);
                    } else if (string.Equals(style.Type, "table", StringComparison.Ordinal) && !TableStyles.ContainsKey(styleId!)) {
                        RecordIndexedStyle();
                        TableStyles.Add(styleId!, style);
                        if (defaultTableStyleId == null && style.IsDefault) {
                            defaultTableStyleId = styleId;
                        }
                    }
                }

                DefaultParagraphStyleId = defaultParagraphStyleId;
                DefaultTableStyleId = defaultTableStyleId;
            }

            internal void RecordStyleChainReference(int depth) {
                if (depth > MaxNativeStyleChainDepth) {
                    throw new System.IO.InvalidDataException($"Word style inheritance depth exceeds the native PDF limit of {MaxNativeStyleChainDepth}.");
                }

                if (_cachedStyleReferenceCount >= MaxNativeCachedStyleReferences) {
                    throw new System.IO.InvalidDataException($"Word style inheritance work exceeds the native PDF limit of {MaxNativeCachedStyleReferences} cached references.");
                }

                _cachedStyleReferenceCount++;
            }

            private void RecordIndexedStyle() {
                if (_indexedStyleCount >= MaxNativeIndexedStyles) {
                    throw new System.IO.InvalidDataException($"Word style count exceeds the native PDF limit of {MaxNativeIndexedStyles}.");
                }

                _indexedStyleCount++;
            }

            internal string? ResolveParagraphStyleId(string? styleId) =>
                string.IsNullOrWhiteSpace(styleId) ? DefaultParagraphStyleId : styleId;

            internal string? ResolveTableStyleId(string? styleId) =>
                string.IsNullOrWhiteSpace(styleId) ? DefaultTableStyleId : styleId;
        }

        private static NativeStyleLookupCache? GetNativeStyleLookupCache(WordDocument? document) {
            if (document == null) {
                return null;
            }

            return NativeStyleLookupCaches.GetValue(
                document,
                current => new NativeStyleLookupCache(WordStyleCatalog.For(current)));
        }

        private static void ResetNativeStyleLookupCache(WordDocument document) =>
            NativeStyleLookupCaches.Remove(document);

        // Matches the style type text case-insensitively, as the loaded-DOM type checks read it.
        private static bool IsNativeStyleType(WordStyleDefinitionsReader.StyleDefinition style, string type) =>
            string.Equals(style.Type, type, StringComparison.OrdinalIgnoreCase);

        /// <summary>Styles of one type by identifier; a style's properties are parsed only when a lookup returns it.</summary>
        private sealed class NativeStyleIndex {
            private readonly Dictionary<string, WordStyleDefinitionsReader.StyleDefinition> _styles = new(StringComparer.Ordinal);

            internal bool ContainsKey(string styleId) => _styles.ContainsKey(styleId);

            internal void Add(string styleId, WordStyleDefinitionsReader.StyleDefinition style) => _styles.Add(styleId, style);

            internal bool TryGetValue(string styleId, out W.Style style) {
                if (_styles.TryGetValue(styleId, out WordStyleDefinitionsReader.StyleDefinition? definition)) {
                    style = definition.Element;
                    return true;
                }
                style = null!;
                return false;
            }
        }
    }
}
