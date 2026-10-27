using System.Collections.Generic;
using System.Runtime.CompilerServices;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private const int MaxNativeIndexedStyles = 4_096;
        private const int MaxNativeStyleChainDepth = 256;
        private const int MaxNativeCachedStyleReferences = 8_192;
        private static readonly ConditionalWeakTable<WordDocument, NativeStyleLookupCache> NativeStyleLookupCaches = new();

        private sealed class NativeStyleLookupCache {
            internal Dictionary<string, W.Style> ParagraphStyles { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, W.Style> CharacterStyles { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, W.Style> TableStyles { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, IReadOnlyList<W.Style>> ParagraphChains { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, IReadOnlyList<W.Style>> CharacterChains { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, IReadOnlyList<W.Style>> TableChains { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, NativeParagraphStyleDefaults> ParagraphDefaults { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, NativeCharacterStyleDefaults> CharacterDefaults { get; } = new(StringComparer.Ordinal);
            internal string? DefaultParagraphStyleId { get; }
            internal string? DefaultTableStyleId { get; }
            private int _indexedStyleCount;
            private int _cachedStyleReferenceCount;

            internal NativeStyleLookupCache(W.Styles? styles) {
                if (styles == null) {
                    return;
                }

                string? defaultParagraphStyleId = null;
                string? defaultTableStyleId = null;
                foreach (W.Style style in styles.Elements<W.Style>()) {
                    string? styleId = style.StyleId?.Value;
                    if (string.IsNullOrEmpty(styleId)) {
                        continue;
                    }

                    if (IsNativeParagraphStyle(style)) {
                        if (!ParagraphStyles.ContainsKey(styleId!)) {
                            RecordIndexedStyle();
                            ParagraphStyles.Add(styleId!, style);
                            if (defaultParagraphStyleId == null && style.Default?.Value == true) {
                                defaultParagraphStyleId = styleId;
                            }
                        }
                    } else if (IsNativeCharacterStyle(style) && !CharacterStyles.ContainsKey(styleId!)) {
                        RecordIndexedStyle();
                        CharacterStyles.Add(styleId!, style);
                    } else if (style.Type?.Value == W.StyleValues.Table && !TableStyles.ContainsKey(styleId!)) {
                        RecordIndexedStyle();
                        TableStyles.Add(styleId!, style);
                        if (defaultTableStyleId == null && style.Default?.Value == true) {
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
                current => new NativeStyleLookupCache(current._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles));
        }

        private static void ResetNativeStyleLookupCache(WordDocument document) =>
            NativeStyleLookupCaches.Remove(document);
    }
}
