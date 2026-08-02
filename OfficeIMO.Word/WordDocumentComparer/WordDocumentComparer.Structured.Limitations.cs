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
            ComparisonWorkBudget comparisonWorkBudget,
            ParagraphNumberingStyleCatalogCache numberingStyleCatalogs) {
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
                    ContainsNumberingFormatting(sourcePart, comparisonWorkBudget, numberingStyleCatalogs),
                    ContainsNumberingFormatting(targetPart, comparisonWorkBudget, numberingStyleCatalogs));
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
            ShapeScanResult sourceScan,
            ShapeScanResult targetScan) {
            if (sourceScan == ShapeScanResult.ResourceLimitExceeded || targetScan == ShapeScanResult.ResourceLimitExceeded) {
                result.AddLimitation(new WordComparisonLimitation(
                    code + ".ResourceLimit",
                    "The bounded comparison-disclosure scan could not determine this shape for every input; increase or simplify the document before relying on shape-presence evidence.",
                    false,
                    false));
            }
            bool sourceContainsShape = sourceScan == ShapeScanResult.Present;
            bool targetContainsShape = targetScan == ShapeScanResult.Present;
            if (!sourceContainsShape && !targetContainsShape) return;
            result.AddLimitation(new WordComparisonLimitation(code, message, sourceContainsShape, targetContainsShape));
        }

        private enum ShapeScanResult {
            Absent,
            Present,
            ResourceLimitExceeded
        }

        private static ShapeScanResult ContainsThemeFormatting(MainDocumentPart mainPart, ComparisonWorkBudget comparisonWorkBudget) {
            var usedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool hasRunWithoutStyle = false;
            bool hasParagraphWithoutStyle = false;
            bool hasTableWithoutStyle = false;
            ShapeScanResult documentScan = ScanComparisonElements(mainPart, comparisonWorkBudget, element => {
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
            });
            if (documentScan != ShapeScanResult.Absent) return documentScan;
            ShapeScanResult defaultStylesScan = ContainsThemeAttribute(
                mainPart.StyleDefinitionsPart?.Styles?.DocDefaults, comparisonWorkBudget);
            if (defaultStylesScan != ShapeScanResult.Absent) return defaultStylesScan;
            defaultStylesScan = ContainsThemeAttribute(
                mainPart.StylesWithEffectsPart?.Styles?.DocDefaults, comparisonWorkBudget);
            if (defaultStylesScan != ShapeScanResult.Absent) return defaultStylesScan;

            IReadOnlyList<Style> styles = CollectComparisonStyles(mainPart, comparisonWorkBudget, out bool styleBudgetExhausted);
            if (styleBudgetExhausted) return ShapeScanResult.ResourceLimitExceeded;
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

        private static ShapeScanResult ContainsThemeAttribute(OpenXmlElement? element, ComparisonWorkBudget comparisonWorkBudget) =>
            element == null
                ? ShapeScanResult.Absent
                : ScanElementSubtree(element, comparisonWorkBudget, candidate => candidate.GetAttributes().Any(IsThemeAttribute));

        private static ShapeScanResult ContainsConditionalTableStyleFormatting(MainDocumentPart mainPart, ComparisonWorkBudget comparisonWorkBudget) {
            var usedTableStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool hasTableWithoutStyle = false;
            ShapeScanResult documentScan = ScanComparisonElements(mainPart, comparisonWorkBudget, element => {
                if (element is TableStyleProperties) return true;
                if (element is TableStyle tableStyle && !string.IsNullOrWhiteSpace(tableStyle.Val?.Value)) {
                    usedTableStyleIds.Add(tableStyle.Val!.Value!);
                }
                if (element is Table table && table.TableProperties?.TableStyle?.Val?.Value == null) hasTableWithoutStyle = true;
                return false;
            });
            if (documentScan != ShapeScanResult.Absent) return documentScan;

            IReadOnlyList<Style> styles = CollectComparisonStyles(mainPart, comparisonWorkBudget, out bool styleBudgetExhausted);
            if (styleBudgetExhausted) return ShapeScanResult.ResourceLimitExceeded;
            if (hasTableWithoutStyle) {
                AddDefaultStyleIds(styles, StyleValues.Table, usedTableStyleIds);
            }
            if (usedTableStyleIds.Count == 0) return ShapeScanResult.Absent;

            ILookup<string, Style> stylesById = styles
                .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
                .ToLookup(style => style.StyleId!.Value!, StringComparer.OrdinalIgnoreCase);
            return ContainsStyleShape(
                usedTableStyleIds,
                stylesById,
                (style, budget) => ScanElementSubtree(style, budget, element => element is TableStyleProperties),
                comparisonWorkBudget);
        }

        private static ShapeScanResult ContainsStyleShape(
            IEnumerable<string> initialStyleIds,
            ILookup<string, Style> stylesById,
            Func<Style, ComparisonWorkBudget, ShapeScanResult> containsShape,
            ComparisonWorkBudget comparisonWorkBudget) {
            var pendingStyleIds = new Stack<string>(initialStyleIds);
            var inspectedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (pendingStyleIds.Count > 0) {
                string styleId = pendingStyleIds.Pop();
                if (!inspectedStyleIds.Add(styleId)) continue;

                foreach (Style style in stylesById[styleId]) {
                    if (!comparisonWorkBudget.TryConsume(1)) return ShapeScanResult.ResourceLimitExceeded;
                    ShapeScanResult styleScan = containsShape(style, comparisonWorkBudget);
                    if (styleScan != ShapeScanResult.Absent) return styleScan;
                    string? baseStyleId = style.BasedOn?.Val?.Value;
                    if (!string.IsNullOrWhiteSpace(baseStyleId)) pendingStyleIds.Push(baseStyleId!);
                }
            }
            return ShapeScanResult.Absent;
        }

        private static ShapeScanResult ContainsNumberingFormatting(
            MainDocumentPart mainPart,
            ComparisonWorkBudget comparisonWorkBudget,
            ParagraphNumberingStyleCatalogCache numberingStyleCatalogs) {
            ParagraphNumberingStyleCatalog styleCatalog = numberingStyleCatalogs.GetOrCreate(mainPart);
            return ScanComparisonElements(mainPart, comparisonWorkBudget, element =>
                element is Paragraph paragraph && ResolveParagraphNumberingProperties(paragraph, styleCatalog) != null);
        }

        private static ShapeScanResult ContainsMoveMarkup(MainDocumentPart mainPart, ComparisonWorkBudget comparisonWorkBudget) =>
            ScanComparisonElements(mainPart, comparisonWorkBudget, element =>
                element.LocalName.StartsWith("moveFrom", StringComparison.Ordinal) ||
                element.LocalName.StartsWith("moveTo", StringComparison.Ordinal));

        private static ShapeScanResult ScanComparisonElements(
            MainDocumentPart mainPart,
            ComparisonWorkBudget comparisonWorkBudget,
            Func<OpenXmlElement, bool> predicate) {
            foreach (OpenXmlCompositeElement root in EnumerateComparisonRoots(mainPart)) {
                if (!comparisonWorkBudget.TryConsume(1)) return ShapeScanResult.ResourceLimitExceeded;
                if (predicate(root)) return ShapeScanResult.Present;
                foreach (OpenXmlElement element in root.Descendants()) {
                    if (!comparisonWorkBudget.TryConsume(1)) return ShapeScanResult.ResourceLimitExceeded;
                    if (predicate(element)) return ShapeScanResult.Present;
                }
            }
            return ShapeScanResult.Absent;
        }

        private static ShapeScanResult ScanElementSubtree(
            OpenXmlElement root,
            ComparisonWorkBudget comparisonWorkBudget,
            Func<OpenXmlElement, bool> predicate) {
            if (!comparisonWorkBudget.TryConsume(1)) return ShapeScanResult.ResourceLimitExceeded;
            if (predicate(root)) return ShapeScanResult.Present;
            foreach (OpenXmlElement element in root.Descendants()) {
                if (!comparisonWorkBudget.TryConsume(1)) return ShapeScanResult.ResourceLimitExceeded;
                if (predicate(element)) return ShapeScanResult.Present;
            }
            return ShapeScanResult.Absent;
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
