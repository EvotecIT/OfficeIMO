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

            if (options.CompareRevisions) {
                AddShapeLimitation(
                    result,
                    "MoveSemantics.RevisionMetadataOnly",
                    "Existing move-from and move-to markup is compared as review metadata; generated redlines use insert/delete revisions and do not synthesize Word move-range semantics.",
                    ContainsMoveMarkup(sourcePart),
                    ContainsMoveMarkup(targetPart));
            }
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
                (ContainsThemeAttribute(mainPart.StyleDefinitionsPart?.Styles?.DocDefaults) ||
                 ContainsThemeAttribute(mainPart.StylesWithEffectsPart?.Styles?.DocDefaults))) return true;

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

            Style[] styles = (mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                .Concat(mainPart.StylesWithEffectsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                .ToArray();
            if (content.OfType<Paragraph>().Any(paragraph => paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == null)) {
                AddDefaultStyleIds(styles, StyleValues.Paragraph, usedStyleIds);
            }
            if (content.OfType<Run>().Any(run => run.RunProperties?.RunStyle?.Val?.Value == null)) {
                AddDefaultStyleIds(styles, StyleValues.Character, usedStyleIds);
            }
            if (content.OfType<Table>().Any(table => table.TableProperties?.TableStyle?.Val?.Value == null)) {
                AddDefaultStyleIds(styles, StyleValues.Table, usedStyleIds);
            }
            ILookup<string, Style> stylesById = styles
                .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                .ToLookup(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase);
            return ContainsStyleShape(
                usedStyleIds,
                stylesById,
                style => new[] { style }.Concat(style.Descendants())
                    .SelectMany(element => element.GetAttributes()).Any(IsThemeAttribute));
        }

        private static void AddDefaultStyleIds(
            IEnumerable<Style> styles,
            StyleValues styleType,
            HashSet<string> styleIds) {
            foreach (Style style in styles.Where(style =>
                         style.Type?.Value == styleType &&
                         style.Default?.Value == true &&
                         !string.IsNullOrWhiteSpace(style.StyleId?.Value))) {
                styleIds.Add(style.StyleId!.Value!);
            }
        }

        private static bool IsThemeAttribute(OpenXmlAttribute attribute) =>
            attribute.LocalName.EndsWith("Theme", StringComparison.OrdinalIgnoreCase) ||
            attribute.LocalName.Equals("themeColor", StringComparison.OrdinalIgnoreCase);

        private static bool ContainsThemeAttribute(OpenXmlElement? element) =>
            element != null &&
            new[] { element }.Concat(element.Descendants())
                .SelectMany(candidate => candidate.GetAttributes())
                .Any(IsThemeAttribute);

        private static bool ContainsConditionalTableStyleFormatting(MainDocumentPart mainPart) {
            OpenXmlElement[] content = EnumerateComparisonRoots(mainPart)
                .SelectMany(root => new[] { root }.Concat(root.Descendants()))
                .ToArray();
            if (content.OfType<TableStyleProperties>().Any()) return true;

            var usedTableStyleIds = new HashSet<string>(content
                .OfType<TableStyle>()
                .Select(tableStyle => tableStyle.Val?.Value)
                .Where(styleId => !string.IsNullOrWhiteSpace(styleId))
                .Select(styleId => styleId!), StringComparer.OrdinalIgnoreCase);
            Style[] styles = (mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                .Concat(mainPart.StylesWithEffectsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                .ToArray();
            if (content.OfType<Table>().Any(table => table.TableProperties?.TableStyle?.Val?.Value == null)) {
                AddDefaultStyleIds(styles, StyleValues.Table, usedTableStyleIds);
            }
            if (usedTableStyleIds.Count == 0) return false;

            ILookup<string, Style> stylesById = styles
                .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                .ToLookup(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase);
            return ContainsStyleShape(
                usedTableStyleIds,
                stylesById,
                style => style.Elements<TableStyleProperties>().Any());
        }

        private static bool ContainsStyleShape(
            IEnumerable<string> initialStyleIds,
            ILookup<string, Style> stylesById,
            Func<Style, bool> containsShape) {
            var pendingStyleIds = new Stack<string>(initialStyleIds);
            var inspectedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (pendingStyleIds.Count > 0) {
                string styleId = pendingStyleIds.Pop();
                if (!inspectedStyleIds.Add(styleId)) continue;

                foreach (Style style in stylesById[styleId]) {
                    if (containsShape(style)) return true;
                    string? baseStyleId = style.BasedOn?.Val?.Value;
                    if (!string.IsNullOrWhiteSpace(baseStyleId)) pendingStyleIds.Push(baseStyleId!);
                }
            }
            return false;
        }

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
