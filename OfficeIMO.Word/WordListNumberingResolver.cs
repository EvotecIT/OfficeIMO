using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Resolves the effective numbering attached to a paragraph, including numbering inherited
/// through paragraph styles and direct <c>w:numId="0"</c> cancellation.
/// </summary>
internal static class WordListNumberingResolver {
    internal sealed class StyleCatalog {
        internal StyleCatalog(MainDocumentPart? mainPart) {
            IEnumerable<Style> styles =
                (mainPart?.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                .Concat(mainPart?.StylesWithEffectsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>());
            ById = styles
                .Where(style => style.Type?.Value == StyleValues.Paragraph && !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                .GroupBy(style => style.StyleId!.Value!, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
            DefaultStyleId = ById.Values.FirstOrDefault(style => style.Default?.Value == true)?.StyleId?.Value;
            LinkedLevels = new Dictionary<(int NumberId, string StyleId), int>();
            Numbering? numbering = mainPart?.NumberingDefinitionsPart?.Numbering;
            if (numbering == null) return;
            Dictionary<int, AbstractNum> abstracts = numbering.Elements<AbstractNum>()
                .Where(abstractNum => abstractNum.AbstractNumberId?.Value != null)
                .GroupBy(abstractNum => abstractNum.AbstractNumberId!.Value)
                .ToDictionary(group => group.Key, group => group.First());
            foreach (NumberingInstance instance in numbering.Elements<NumberingInstance>()) {
                if (instance.NumberID?.Value is not int numberId ||
                    instance.AbstractNumId?.Val?.Value is not int abstractId ||
                    !abstracts.TryGetValue(abstractId, out AbstractNum? abstractNum)) continue;
                Dictionary<int, Level> overrides = instance.Elements<LevelOverride>()
                    .Where(levelOverride => levelOverride.LevelIndex?.Value != null && levelOverride.GetFirstChild<Level>() != null)
                    .GroupBy(levelOverride => levelOverride.LevelIndex!.Value)
                    .ToDictionary(group => group.Key, group => group.First().GetFirstChild<Level>()!);
                foreach (Level abstractLevel in abstractNum.Elements<Level>()) {
                    if (abstractLevel.LevelIndex?.Value is not int index) continue;
                    Level level = overrides.TryGetValue(index, out Level? replacement) ? replacement : abstractLevel;
                    if (level.GetFirstChild<ParagraphStyleIdInLevel>()?.Val?.Value is string linkedStyle) {
                        LinkedLevels[(numberId, linkedStyle)] = index;
                    }
                }
                foreach (KeyValuePair<int, Level> levelOverride in overrides) {
                    if (levelOverride.Value.GetFirstChild<ParagraphStyleIdInLevel>()?.Val?.Value is string linkedStyle) {
                        LinkedLevels[(numberId, linkedStyle)] = levelOverride.Key;
                    }
                }
            }
        }

        internal Dictionary<string, Style> ById { get; }
        internal string? DefaultStyleId { get; }
        internal Dictionary<(int NumberId, string StyleId), int> LinkedLevels { get; }
    }

    internal static StyleCatalog CreateStyleCatalog(WordDocument document) =>
        new(document._wordprocessingDocument.MainDocumentPart);

    internal readonly struct ResolvedNumbering {
        internal ResolvedNumbering(int numberId, int level) {
            NumberId = numberId;
            Level = level;
        }

        internal int NumberId { get; }
        internal int Level { get; }
    }

    internal static bool TryResolve(WordParagraph paragraph, out ResolvedNumbering numbering, StyleCatalog? styleCatalog = null) {
        numbering = default;
        if (paragraph?._paragraph == null || paragraph._document == null) {
            return false;
        }

        NumberingProperties? direct = paragraph._paragraph.ParagraphProperties?.NumberingProperties;
        int? numberId = ReadNumberId(direct);
        int? level = ReadLevel(direct);
        if (numberId == 0) {
            return false;
        }

        // The common direct-numbering path needs no style catalog lookup.
        if (numberId > 0 && level.HasValue) {
            numbering = new ResolvedNumbering(numberId.Value, Math.Max(0, level.Value));
            return true;
        }

        if (!numberId.HasValue || !level.HasValue) {
            styleCatalog ??= CreateStyleCatalog(paragraph._document);
            string? styleId = paragraph._paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            NumberingProperties? inherited = ResolveStyleNumbering(styleCatalog, styleId);
            numberId ??= ReadNumberId(inherited);
            if (!level.HasValue && numberId > 0) {
                level = ResolveLinkedLevel(styleCatalog, numberId.Value, styleId) ?? ReadLevel(inherited);
            }
        }

        if (!numberId.HasValue || numberId.Value <= 0) {
            return false;
        }

        numbering = new ResolvedNumbering(numberId.Value, Math.Max(0, level ?? 0));
        return true;
    }

    private static NumberingProperties? ResolveStyleNumbering(StyleCatalog catalog, string? styleId) {
        if (string.IsNullOrWhiteSpace(styleId)) styleId = catalog.DefaultStyleId;

        int? numberId = null;
        int? level = null;
        string? currentStyleId = styleId;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!)) {
            if (!catalog.ById.TryGetValue(currentStyleId!, out Style? style)) {
                break;
            }

            NumberingProperties? candidate = style.StyleParagraphProperties?.NumberingProperties;
            numberId ??= ReadNumberId(candidate);
            level ??= ReadLevel(candidate);
            if (numberId.HasValue && level.HasValue) {
                break;
            }

            currentStyleId = style.BasedOn?.Val?.Value;
        }

        if (!numberId.HasValue && !level.HasValue) {
            return null;
        }

        var resolved = new NumberingProperties();
        if (level.HasValue) {
            resolved.Append(new NumberingLevelReference { Val = level.Value });
        }
        if (numberId.HasValue) {
            resolved.Append(new NumberingId { Val = numberId.Value });
        }
        return resolved;
    }

    private static int? ResolveLinkedLevel(StyleCatalog catalog, int numberId, string? styleId) {
        string? currentStyleId = string.IsNullOrWhiteSpace(styleId) ? catalog.DefaultStyleId : styleId;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!)) {
            if (catalog.LinkedLevels.TryGetValue((numberId, currentStyleId!), out int level)) return level;
            currentStyleId = catalog.ById.TryGetValue(currentStyleId!, out Style? style)
                ? style.BasedOn?.Val?.Value
                : null;
        }
        return null;
    }

    private static int? ReadNumberId(NumberingProperties? properties) =>
        properties?.NumberingId?.Val?.Value;

    private static int? ReadLevel(NumberingProperties? properties) =>
        properties?.NumberingLevelReference?.Val?.Value;
}
