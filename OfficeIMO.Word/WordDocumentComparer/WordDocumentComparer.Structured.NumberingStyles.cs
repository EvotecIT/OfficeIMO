using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static NumberingProperties? ResolveParagraphNumberingProperties(
            Paragraph paragraph,
            ParagraphNumberingStyleCatalog styleCatalog) {
            NumberingProperties? directNumbering = paragraph.ParagraphProperties?.NumberingProperties;
            if (directNumbering != null) return directNumbering;

            var pendingStyleIds = new Stack<string>();
            string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(styleId)) {
                pendingStyleIds.Push(styleId!);
            } else {
                foreach (string defaultStyleId in styleCatalog.DefaultParagraphStyleIds) {
                    pendingStyleIds.Push(defaultStyleId);
                }
            }

            var visited = new HashSet<string>(StringComparer.Ordinal);
            while (pendingStyleIds.Count > 0) {
                string currentStyleId = pendingStyleIds.Pop();
                if (!visited.Add(currentStyleId)) continue;
                foreach (Style style in styleCatalog.StylesById[currentStyleId]) {
                    NumberingProperties? numbering = style.StyleParagraphProperties?.NumberingProperties;
                    if (numbering != null) return numbering;
                    string? baseStyleId = style.BasedOn?.Val?.Value;
                    if (!string.IsNullOrWhiteSpace(baseStyleId)) pendingStyleIds.Push(baseStyleId!);
                }
            }

            return null;
        }

        /// <summary>Caches the combined style definitions used while resolving paragraph numbering for one analysis.</summary>
        private sealed class ParagraphNumberingStyleCatalog {
            private ParagraphNumberingStyleCatalog(
                ILookup<string, Style> stylesById,
                IReadOnlyList<string> defaultParagraphStyleIds) {
                StylesById = stylesById;
                DefaultParagraphStyleIds = defaultParagraphStyleIds;
            }

            internal ILookup<string, Style> StylesById { get; }
            internal IReadOnlyList<string> DefaultParagraphStyleIds { get; }

            internal static ParagraphNumberingStyleCatalog Create(MainDocumentPart mainPart) {
                Style[] styles = (mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                    .Concat(mainPart.StylesWithEffectsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                    .ToArray();
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
                catalog = ParagraphNumberingStyleCatalog.Create(mainPart);
                _catalogs.Add(mainPart, catalog);
                return catalog;
            }
        }
    }
}
