using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static NumberingProperties? ResolveParagraphNumberingProperties(
            Paragraph paragraph,
            ParagraphNumberingStyleCatalog styleCatalog) {
            NumberingProperties? directNumbering = paragraph.ParagraphProperties?.NumberingProperties;
            if (directNumbering != null) return directNumbering;
            string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            return styleCatalog.ResolveNumbering(styleId);
        }

        /// <summary>Caches the combined style definitions used while resolving paragraph numbering for one analysis.</summary>
        private sealed class ParagraphNumberingStyleCatalog {
            private const string DefaultParagraphStylesCacheKey = "\0default-paragraph-styles";
            private readonly Dictionary<string, NumberingProperties?> _resolvedNumbering =
                new(StringComparer.Ordinal);

            private ParagraphNumberingStyleCatalog(
                ILookup<string, Style> stylesById,
                IReadOnlyList<string> defaultParagraphStyleIds) {
                StylesById = stylesById;
                DefaultParagraphStyleIds = defaultParagraphStyleIds;
            }

            internal ILookup<string, Style> StylesById { get; }
            internal IReadOnlyList<string> DefaultParagraphStyleIds { get; }

            internal NumberingProperties? ResolveNumbering(string? styleId) {
                string cacheKey = string.IsNullOrWhiteSpace(styleId)
                    ? DefaultParagraphStylesCacheKey
                    : styleId!;
                if (_resolvedNumbering.TryGetValue(cacheKey, out NumberingProperties? cached)) return cached;

                var pendingStyleIds = new Stack<string>();
                if (cacheKey == DefaultParagraphStylesCacheKey) {
                    foreach (string defaultStyleId in DefaultParagraphStyleIds) pendingStyleIds.Push(defaultStyleId);
                } else {
                    pendingStyleIds.Push(cacheKey);
                }
                var visited = new HashSet<string>(StringComparer.Ordinal);
                while (pendingStyleIds.Count > 0) {
                    string currentStyleId = pendingStyleIds.Pop();
                    if (!visited.Add(currentStyleId)) continue;
                    foreach (Style style in StylesById[currentStyleId]) {
                        NumberingProperties? numbering = style.StyleParagraphProperties?.NumberingProperties;
                        if (numbering != null) {
                            _resolvedNumbering[cacheKey] = numbering;
                            return numbering;
                        }
                        string? baseStyleId = style.BasedOn?.Val?.Value;
                        if (!string.IsNullOrWhiteSpace(baseStyleId)) pendingStyleIds.Push(baseStyleId!);
                    }
                }
                _resolvedNumbering[cacheKey] = null;
                return null;
            }

            internal static ParagraphNumberingStyleCatalog? Create(
                MainDocumentPart mainPart,
                ComparisonWorkBudget? comparisonWorkBudget) {
                var styles = new List<Style>();
                IEnumerable<Style> candidates =
                    (mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                    .Concat(mainPart.StylesWithEffectsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>());
                foreach (Style style in candidates) {
                    if (comparisonWorkBudget != null && !comparisonWorkBudget.TryConsume(1)) return null;
                    styles.Add(style);
                }
                ILookup<string, Style> stylesById = styles
                    .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                    .ToLookup(style => style.StyleId!.Value!, StringComparer.Ordinal);
                string[] defaultParagraphStyleIds = styles
                    .Where(style =>
                        style.Type?.Value == StyleValues.Paragraph &&
                        style.Default?.Value == true &&
                        !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                    .Select(style => style.StyleId!.Value!)
                    .ToArray();
                return new ParagraphNumberingStyleCatalog(stylesById, defaultParagraphStyleIds);
            }
        }

        /// <summary>Reuses numbering style catalogs across all feature and limitation passes in one comparison.</summary>
        private sealed class ParagraphNumberingStyleCatalogCache {
            private readonly Dictionary<MainDocumentPart, ParagraphNumberingStyleCatalog> _catalogs = new();

            internal ParagraphNumberingStyleCatalog GetOrCreate(MainDocumentPart mainPart) {
                if (_catalogs.TryGetValue(mainPart, out ParagraphNumberingStyleCatalog? catalog)) return catalog;
                catalog = ParagraphNumberingStyleCatalog.Create(mainPart, comparisonWorkBudget: null)!;
                _catalogs.Add(mainPart, catalog);
                return catalog;
            }

            internal bool TryGetOrCreate(
                MainDocumentPart mainPart,
                ComparisonWorkBudget comparisonWorkBudget,
                out ParagraphNumberingStyleCatalog? catalog) {
                if (_catalogs.TryGetValue(mainPart, out catalog)) return true;
                catalog = ParagraphNumberingStyleCatalog.Create(mainPart, comparisonWorkBudget);
                if (catalog == null) return false;
                _catalogs.Add(mainPart, catalog);
                return true;
            }
        }
    }
}
