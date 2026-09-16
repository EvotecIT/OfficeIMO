using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// Resolves the effective numbering attached to a paragraph, including numbering inherited
/// through paragraph styles and direct <c>w:numId="0"</c> cancellation.
/// </summary>
internal static class WordListNumberingResolver {
    internal readonly struct ResolvedNumbering {
        internal ResolvedNumbering(int numberId, int level) {
            NumberId = numberId;
            Level = level;
        }

        internal int NumberId { get; }
        internal int Level { get; }
    }

    internal static bool TryResolve(WordParagraph paragraph, out ResolvedNumbering numbering) {
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
            NumberingProperties? inherited = ResolveStyleNumbering(
                paragraph._document._wordprocessingDocument.MainDocumentPart,
                paragraph._paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            numberId ??= ReadNumberId(inherited);
            level ??= ReadLevel(inherited);
        }

        if (!numberId.HasValue || numberId.Value <= 0) {
            return false;
        }

        numbering = new ResolvedNumbering(numberId.Value, Math.Max(0, level ?? 0));
        return true;
    }

    private static NumberingProperties? ResolveStyleNumbering(MainDocumentPart? mainPart, string? styleId) {
        if (mainPart == null) {
            return null;
        }

        IEnumerable<Style> styles =
            (mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
            .Concat(mainPart.StylesWithEffectsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>());
        Dictionary<string, Style> byId = styles
            .Where(style => style.Type?.Value == StyleValues.Paragraph && !string.IsNullOrWhiteSpace(style.StyleId?.Value))
            .GroupBy(style => style.StyleId!.Value!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

        if (string.IsNullOrWhiteSpace(styleId)) {
            styleId = byId.Values.FirstOrDefault(style => style.Default?.Value == true)?.StyleId?.Value;
        }

        int? numberId = null;
        int? level = null;
        string? currentStyleId = styleId;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(currentStyleId) && visited.Add(currentStyleId!)) {
            if (!byId.TryGetValue(currentStyleId!, out Style? style)) {
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

    private static int? ReadNumberId(NumberingProperties? properties) =>
        properties?.NumberingId?.Val?.Value;

    private static int? ReadLevel(NumberingProperties? properties) =>
        properties?.NumberingLevelReference?.Val?.Value;
}
