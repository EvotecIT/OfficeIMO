using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeKnownComparisonLimitations(
            WordDocument source,
            WordDocument target,
            WordComparisonResult result,
            WordComparisonOptions options,
            ComparisonWorkBudget comparisonWorkBudget) {
            MainDocumentPart? sourcePart = source._wordprocessingDocument.MainDocumentPart;
            MainDocumentPart? targetPart = target._wordprocessingDocument.MainDocumentPart;
            if (sourcePart == null || targetPart == null) return;

            if (options.CompareEffectiveFormatting) {
                AddShapeLimitation(
                    result,
                    "EffectiveFormatting.ThemeResolution",
                    "Theme font and theme color tokens are compared as declared tokens; Office theme resolution and layout-dependent appearance are not evaluated.",
                    ContainsThemeFormatting(sourcePart, comparisonWorkBudget),
                    ContainsThemeFormatting(targetPart, comparisonWorkBudget));
                AddShapeLimitation(
                    result,
                    "EffectiveFormatting.ConditionalTableStyles",
                    "Conditional table-style regions are compared structurally; the full Word table-style cascade is not folded into effective run or paragraph formatting.",
                    ContainsConditionalTableStyleFormatting(sourcePart, comparisonWorkBudget),
                    ContainsConditionalTableStyleFormatting(targetPart, comparisonWorkBudget));
                AddShapeLimitation(
                    result,
                    "EffectiveFormatting.NumberingLevelStyles",
                    "List definitions and paragraph numbering are compared, but numbering-level style inheritance is not folded into effective run or paragraph formatting.",
                    ContainsNumberingFormatting(sourcePart, comparisonWorkBudget),
                    ContainsNumberingFormatting(targetPart, comparisonWorkBudget));
            }

            if (options.CompareRevisions) {
                AddShapeLimitation(
                    result,
                    "MoveSemantics.RevisionMetadataOnly",
                    "Existing move-from and move-to markup is compared as review metadata; generated redlines use insert/delete revisions and do not synthesize Word move-range semantics.",
                    ContainsMoveMarkup(sourcePart, comparisonWorkBudget),
                    ContainsMoveMarkup(targetPart, comparisonWorkBudget));
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

        private static bool ContainsThemeFormatting(MainDocumentPart mainPart, ComparisonWorkBudget comparisonWorkBudget) {
            var usedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool hasRunWithoutStyle = false;
            bool hasParagraphWithoutStyle = false;
            bool hasTableWithoutStyle = false;
            if (ScanComparisonElements(mainPart, comparisonWorkBudget, element => {
                if (element.GetAttributes().Any(IsThemeAttribute)) return true;
                if (element is Run run && run.RunProperties?.RunStyle?.Val?.Value == null) hasRunWithoutStyle = true;
                if (element is Paragraph paragraph && paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value == null) hasParagraphWithoutStyle = true;
                if (element is Table table && table.TableProperties?.TableStyle?.Val?.Value == null) hasTableWithoutStyle = true;
                string? styleId = element switch {
                    ParagraphStyleId paragraphStyle => paragraphStyle.Val?.Value,
                    RunStyle runStyle => runStyle.Val?.Value,
                    TableStyle tableStyle => tableStyle.Val?.Value,
                    _ => null
                };
                if (!string.IsNullOrWhiteSpace(styleId)) usedStyleIds.Add(styleId!);
                return false;
            })) return true;
            if (hasRunWithoutStyle &&
                (ContainsThemeAttribute(mainPart.StyleDefinitionsPart?.Styles?.DocDefaults, comparisonWorkBudget) ||
                 ContainsThemeAttribute(mainPart.StylesWithEffectsPart?.Styles?.DocDefaults, comparisonWorkBudget))) return true;

            IReadOnlyList<Style> styles = CollectComparisonStyles(mainPart, comparisonWorkBudget, out bool styleBudgetExhausted);
            if (styleBudgetExhausted) return true;
            if (hasParagraphWithoutStyle) {
                AddDefaultStyleIds(styles, StyleValues.Paragraph, usedStyleIds);
            }
            if (hasRunWithoutStyle) {
                AddDefaultStyleIds(styles, StyleValues.Character, usedStyleIds);
            }
            if (hasTableWithoutStyle) {
                AddDefaultStyleIds(styles, StyleValues.Table, usedStyleIds);
            }
            ILookup<string, Style> stylesById = styles
                .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                .ToLookup(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase);
            return ContainsStyleShape(
                usedStyleIds,
                stylesById,
                (style, budget) => ScanElementSubtree(style, budget, element => element.GetAttributes().Any(IsThemeAttribute)),
                comparisonWorkBudget);
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

        private static bool ContainsThemeAttribute(OpenXmlElement? element, ComparisonWorkBudget comparisonWorkBudget) =>
            element != null && ScanElementSubtree(element, comparisonWorkBudget, candidate => candidate.GetAttributes().Any(IsThemeAttribute));

        private static bool ContainsConditionalTableStyleFormatting(MainDocumentPart mainPart, ComparisonWorkBudget comparisonWorkBudget) {
            var usedTableStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool hasTableWithoutStyle = false;
            if (ScanComparisonElements(mainPart, comparisonWorkBudget, element => {
                if (element is TableStyleProperties) return true;
                if (element is TableStyle tableStyle && !string.IsNullOrWhiteSpace(tableStyle.Val?.Value)) {
                    usedTableStyleIds.Add(tableStyle.Val!.Value!);
                }
                if (element is Table table && table.TableProperties?.TableStyle?.Val?.Value == null) hasTableWithoutStyle = true;
                return false;
            })) return true;

            IReadOnlyList<Style> styles = CollectComparisonStyles(mainPart, comparisonWorkBudget, out bool styleBudgetExhausted);
            if (styleBudgetExhausted) return true;
            if (hasTableWithoutStyle) {
                AddDefaultStyleIds(styles, StyleValues.Table, usedTableStyleIds);
            }
            if (usedTableStyleIds.Count == 0) return false;

            ILookup<string, Style> stylesById = styles
                .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                .ToLookup(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase);
            return ContainsStyleShape(
                usedTableStyleIds,
                stylesById,
                (style, budget) => ScanElementSubtree(style, budget, element => element is TableStyleProperties),
                comparisonWorkBudget);
        }

        private static bool ContainsStyleShape(
            IEnumerable<string> initialStyleIds,
            ILookup<string, Style> stylesById,
            Func<Style, ComparisonWorkBudget, bool> containsShape,
            ComparisonWorkBudget comparisonWorkBudget) {
            var pendingStyleIds = new Stack<string>(initialStyleIds);
            var inspectedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (pendingStyleIds.Count > 0) {
                string styleId = pendingStyleIds.Pop();
                if (!inspectedStyleIds.Add(styleId)) continue;

                foreach (Style style in stylesById[styleId]) {
                    if (!comparisonWorkBudget.TryConsume(1) || containsShape(style, comparisonWorkBudget)) return true;
                    string? baseStyleId = style.BasedOn?.Val?.Value;
                    if (!string.IsNullOrWhiteSpace(baseStyleId)) pendingStyleIds.Push(baseStyleId!);
                }
            }
            return false;
        }

        private static bool ContainsNumberingFormatting(MainDocumentPart mainPart, ComparisonWorkBudget comparisonWorkBudget) =>
            ScanComparisonElements(mainPart, comparisonWorkBudget, element =>
                element is Paragraph paragraph && ResolveParagraphNumberingProperties(paragraph, mainPart) != null);

        private static bool ContainsMoveMarkup(MainDocumentPart mainPart, ComparisonWorkBudget comparisonWorkBudget) =>
            ScanComparisonElements(mainPart, comparisonWorkBudget, element =>
                element.LocalName.StartsWith("moveFrom", StringComparison.Ordinal) ||
                element.LocalName.StartsWith("moveTo", StringComparison.Ordinal));

        private static bool ScanComparisonElements(
            MainDocumentPart mainPart,
            ComparisonWorkBudget comparisonWorkBudget,
            Func<OpenXmlElement, bool> predicate) {
            foreach (OpenXmlCompositeElement root in EnumerateComparisonRoots(mainPart)) {
                if (!comparisonWorkBudget.TryConsume(1) || predicate(root)) return true;
                foreach (OpenXmlElement element in root.Descendants()) {
                    if (!comparisonWorkBudget.TryConsume(1) || predicate(element)) return true;
                }
            }
            return false;
        }

        private static bool ScanElementSubtree(
            OpenXmlElement root,
            ComparisonWorkBudget comparisonWorkBudget,
            Func<OpenXmlElement, bool> predicate) {
            if (!comparisonWorkBudget.TryConsume(1) || predicate(root)) return true;
            foreach (OpenXmlElement element in root.Descendants()) {
                if (!comparisonWorkBudget.TryConsume(1) || predicate(element)) return true;
            }
            return false;
        }

        private static IReadOnlyList<Style> CollectComparisonStyles(
            MainDocumentPart mainPart,
            ComparisonWorkBudget comparisonWorkBudget,
            out bool budgetExhausted) {
            var styles = new List<Style>();
            IEnumerable<Style> candidates = (mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>())
                .Concat(mainPart.StylesWithEffectsPart?.Styles?.Elements<Style>() ?? Enumerable.Empty<Style>());
            foreach (Style style in candidates) {
                if (!comparisonWorkBudget.TryConsume(1)) {
                    budgetExhausted = true;
                    return styles;
                }
                styles.Add(style);
            }
            budgetExhausted = false;
            return styles;
        }

        private static IEnumerable<OpenXmlCompositeElement> EnumerateComparisonRoots(MainDocumentPart mainPart) =>
            WordFieldInventory.EnumerateFieldRoots(mainPart).Select(root => root.Root);
    }
}
