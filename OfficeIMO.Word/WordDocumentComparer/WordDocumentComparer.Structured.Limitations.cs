using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeKnownComparisonLimitations(
            WordDocument source,
            WordDocument target,
            WordComparisonResult result,
            WordComparisonOptions options) {
            MainDocumentPart? sourcePart = source._wordprocessingDocument.MainDocumentPart;
            MainDocumentPart? targetPart = target._wordprocessingDocument.MainDocumentPart;
            if (sourcePart == null || targetPart == null) return;

            if (options.CompareEffectiveFormatting) {
                AddShapeLimitation(
                    result,
                    "EffectiveFormatting.ThemeResolution",
                    "Theme font and theme color tokens are compared as declared tokens; Office theme resolution and layout-dependent appearance are not evaluated.",
                    ContainsThemeFormatting(sourcePart),
                    ContainsThemeFormatting(targetPart));
                AddShapeLimitation(
                    result,
                    "EffectiveFormatting.ConditionalTableStyles",
                    "Conditional table-style regions are compared structurally; the full Word table-style cascade is not folded into effective run or paragraph formatting.",
                    ContainsConditionalTableStyleFormatting(sourcePart),
                    ContainsConditionalTableStyleFormatting(targetPart));
                AddShapeLimitation(
                    result,
                    "EffectiveFormatting.NumberingLevelStyles",
                    "List definitions and paragraph numbering are compared, but numbering-level style inheritance is not folded into effective run or paragraph formatting.",
                    ContainsNumberingFormatting(sourcePart),
                    ContainsNumberingFormatting(targetPart));
            }

            AddShapeLimitation(
                result,
                "MoveSemantics.RevisionMetadataOnly",
                "Existing move-from and move-to markup is compared as review metadata; generated redlines use insert/delete revisions and do not synthesize Word move-range semantics.",
                ContainsMoveMarkup(sourcePart),
                ContainsMoveMarkup(targetPart));
        }

        private static void AddShapeLimitation(
            WordComparisonResult result,
            string code,
            string message,
            bool sourceContainsShape,
            bool targetContainsShape) {
            if (!sourceContainsShape && !targetContainsShape) return;
            result.AddLimitation(new WordComparisonLimitation(code, message, sourceContainsShape, targetContainsShape));
        }

        private static bool ContainsThemeFormatting(MainDocumentPart mainPart) {
            OpenXmlElement[] content = EnumerateComparisonRoots(mainPart)
                .SelectMany(root => new[] { root }.Concat(root.Descendants()))
                .ToArray();
            if (content.SelectMany(element => element.GetAttributes()).Any(IsThemeAttribute)) return true;
            if (content.OfType<Run>().Any() &&
                ContainsThemeAttribute(mainPart.StyleDefinitionsPart?.Styles?.DocDefaults)) return true;

            var usedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (OpenXmlElement element in content) {
                string? styleId = element switch {
                    ParagraphStyleId paragraphStyle => paragraphStyle.Val?.Value,
                    RunStyle runStyle => runStyle.Val?.Value,
                    TableStyle tableStyle => tableStyle.Val?.Value,
                    _ => null
                };
                if (!string.IsNullOrWhiteSpace(styleId)) usedStyleIds.Add(styleId!);
            }

            IEnumerable<Style> styles = (mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                .Concat(mainPart.StylesWithEffectsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>());
            Dictionary<string, Style> stylesById = styles
                .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                .GroupBy(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            var inspectedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string usedStyleId in usedStyleIds) {
                string? currentStyleId = usedStyleId;
                while (!string.IsNullOrWhiteSpace(currentStyleId) && inspectedStyleIds.Add(currentStyleId!)) {
                    if (!stylesById.TryGetValue(currentStyleId!, out Style? style)) break;
                    if (new[] { style }.Concat(style.Descendants())
                        .SelectMany(element => element.GetAttributes()).Any(IsThemeAttribute)) return true;
                    currentStyleId = style.BasedOn?.Val?.Value;
                }
            }
            return false;
        }

        private static bool IsThemeAttribute(OpenXmlAttribute attribute) =>
            attribute.LocalName.EndsWith("Theme", StringComparison.OrdinalIgnoreCase) ||
            attribute.LocalName.Equals("themeColor", StringComparison.OrdinalIgnoreCase);

        private static bool ContainsThemeAttribute(OpenXmlElement? element) =>
            element != null &&
            new[] { element }.Concat(element.Descendants())
                .SelectMany(candidate => candidate.GetAttributes())
                .Any(IsThemeAttribute);

        private static bool ContainsConditionalTableStyleFormatting(MainDocumentPart mainPart) =>
            EnumerateComparisonRoots(mainPart).Any(root =>
                root.Descendants<TableStyle>().Any() ||
                root.Descendants<TableStyleConditionalFormattingTableRowProperties>().Any() ||
                root.Descendants<TableStyleConditionalFormattingTableCellProperties>().Any());

        private static bool ContainsNumberingFormatting(MainDocumentPart mainPart) =>
            EnumerateComparisonRoots(mainPart)
                .SelectMany(root => new[] { root }.Concat(root.Descendants()))
                .OfType<Paragraph>()
                .Any(paragraph => ResolveParagraphNumberingProperties(paragraph, mainPart) != null);

        private static bool ContainsMoveMarkup(MainDocumentPart mainPart) =>
            EnumerateComparisonRoots(mainPart)
                .SelectMany(root => root.Descendants())
                .Any(element => element.LocalName.StartsWith("moveFrom", StringComparison.Ordinal) ||
                                element.LocalName.StartsWith("moveTo", StringComparison.Ordinal));

        private static IEnumerable<OpenXmlCompositeElement> EnumerateComparisonRoots(MainDocumentPart mainPart) =>
            WordFieldInventory.EnumerateFieldRoots(mainPart).Select(root => root.Root);
    }
}
