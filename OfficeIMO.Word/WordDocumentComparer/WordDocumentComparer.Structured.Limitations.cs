using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static void AnalyzeKnownComparisonLimitations(
            WordPackageView source,
            WordPackageView target,
            WordComparisonResult result,
            WordComparisonOptions options,
            ComparisonWorkBudget comparisonWorkBudget,
            ParagraphNumberingStyleCatalogCache numberingStyleCatalogs) {
            WordPartView? sourcePart = source.Main;
            WordPartView? targetPart = target.Main;
            if (sourcePart == null || targetPart == null) return;

            if (options.CompareEffectiveFormatting) {
                AddShapeLimitation(
                    result,
                    "effective-formatting-theme-resolution",
                    "Theme font and theme color tokens are compared as declared tokens; Office theme resolution and layout-dependent appearance are not evaluated.",
                    ContainsThemeFormatting(sourcePart, comparisonWorkBudget),
                    ContainsThemeFormatting(targetPart, comparisonWorkBudget));
                AddShapeLimitation(
                    result,
                    "effective-formatting-conditional-table-styles",
                    "Conditional table-style regions are compared structurally; the full Word table-style cascade is not folded into effective run or paragraph formatting.",
                    ContainsConditionalTableStyleFormatting(sourcePart, comparisonWorkBudget),
                    ContainsConditionalTableStyleFormatting(targetPart, comparisonWorkBudget));
                AddShapeLimitation(
                    result,
                    "effective-formatting-numbering-level-styles",
                    "List definitions and paragraph numbering are compared, but numbering-level style inheritance is not folded into effective run or paragraph formatting.",
                    ContainsNumberingFormatting(sourcePart, comparisonWorkBudget, numberingStyleCatalogs),
                    ContainsNumberingFormatting(targetPart, comparisonWorkBudget, numberingStyleCatalogs));
            }

            if (options.CompareRevisions) {
                AddShapeLimitation(
                    result,
                    "move-semantics-revision-metadata-only",
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
            bool sourceContainsShape = sourceScan == ShapeScanResult.Present;
            bool targetContainsShape = targetScan == ShapeScanResult.Present;
            if (sourceScan == ShapeScanResult.ResourceLimitExceeded || targetScan == ShapeScanResult.ResourceLimitExceeded) {
                result.AddLimitation(new WordComparisonLimitation(
                    code + ".ResourceLimit",
                    "The bounded comparison-disclosure scan could not determine this shape for every input; increase or simplify the document before relying on shape-presence evidence.",
                    sourceContainsShape,
                    targetContainsShape));
                return;
            }
            if (!sourceContainsShape && !targetContainsShape) return;
            result.AddLimitation(new WordComparisonLimitation(code, message, sourceContainsShape, targetContainsShape));
        }

        private enum ShapeScanResult {
            Absent,
            Present,
            ResourceLimitExceeded
        }

        private static ShapeScanResult ContainsThemeFormatting(WordPartView mainPart, ComparisonWorkBudget comparisonWorkBudget) {
            var usedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool hasRunWithoutStyle = false;
            bool hasParagraphWithoutStyle = false;
            bool hasTableWithoutStyle = false;
            ShapeScanResult documentScan = ScanComparisonElements(mainPart, comparisonWorkBudget, element => {
                if (HasThemeAttribute(element)) return true;
                if (element.NamespaceId != W || IsInsideExtensionList(element)) return false;
                if (element.LocalName == "r" && ChildVal(element.Child(W, "rPr"), "rStyle") == null) hasRunWithoutStyle = true;
                if (element.LocalName == "p" && ChildVal(element.Child(W, "pPr"), "pStyle") == null) hasParagraphWithoutStyle = true;
                if (element.LocalName == "tbl" && ChildVal(element.Child(W, "tblPr"), "tblStyle") == null) hasTableWithoutStyle = true;
                string? styleId = element.LocalName is "pStyle" or "rStyle" or "tblStyle" ? WVal(element) : null;
                if (!string.IsNullOrWhiteSpace(styleId)) usedStyleIds.Add(styleId!);
                return false;
            });
            if (documentScan != ShapeScanResult.Absent) return documentScan;
            ShapeScanResult defaultStylesScan = ContainsThemeAttribute(
                GetComparisonStyles(mainPart).DocumentDefaults, comparisonWorkBudget);
            if (defaultStylesScan != ShapeScanResult.Absent) return defaultStylesScan;
            defaultStylesScan = ContainsThemeAttribute(
                GetComparisonStyles(mainPart, WordPackageView.StylesWithEffectsRelationshipType).DocumentDefaults, comparisonWorkBudget);
            if (defaultStylesScan != ShapeScanResult.Absent) return defaultStylesScan;

            IReadOnlyList<ComparisonStyle> styles = CollectComparisonStyles(mainPart, comparisonWorkBudget, out bool styleBudgetExhausted);
            if (styleBudgetExhausted) return ShapeScanResult.ResourceLimitExceeded;
            if (hasParagraphWithoutStyle) {
                AddDefaultStyleIds(styles, "paragraph", usedStyleIds);
            }
            if (hasRunWithoutStyle) {
                AddDefaultStyleIds(styles, "character", usedStyleIds);
            }
            if (hasTableWithoutStyle) {
                AddDefaultStyleIds(styles, "table", usedStyleIds);
            }
            ILookup<string, ComparisonStyle> stylesById = styles
                .Where(style => !string.IsNullOrWhiteSpace(style.StyleId))
                .ToLookup(style => style.StyleId!, StringComparer.OrdinalIgnoreCase);
            return ContainsStyleShape(
                usedStyleIds,
                stylesById,
                (style, budget) => ScanElementSubtree(style, budget, HasThemeAttribute),
                comparisonWorkBudget);
        }

        private static void AddDefaultStyleIds(
            IEnumerable<ComparisonStyle> styles,
            string styleType,
            HashSet<string> styleIds) {
            foreach (ComparisonStyle style in styles.Where(style =>
                         string.Equals(style.Type, styleType, StringComparison.Ordinal) &&
                         style.IsDefault &&
                         !string.IsNullOrWhiteSpace(style.StyleId))) {
                styleIds.Add(style.StyleId!);
            }
        }

        private static bool HasThemeAttribute(OfficeXmlNode element) =>
            WordXmlNodes.AnyAttribute(element, static (localName, _, _) =>
                localName.EndsWith("Theme", StringComparison.OrdinalIgnoreCase) ||
                localName.Equals("themeColor", StringComparison.OrdinalIgnoreCase));

        private static ShapeScanResult ContainsThemeAttribute(OfficeXmlNode element, ComparisonWorkBudget comparisonWorkBudget) =>
            element.IsNone
                ? ShapeScanResult.Absent
                : ScanElementSubtree(element, comparisonWorkBudget, HasThemeAttribute);

        private static ShapeScanResult ContainsConditionalTableStyleFormatting(WordPartView mainPart, ComparisonWorkBudget comparisonWorkBudget) {
            var usedTableStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool hasTableWithoutStyle = false;
            ShapeScanResult documentScan = ScanComparisonElements(mainPart, comparisonWorkBudget, element => {
                if (element.NamespaceId != W || IsInsideExtensionList(element)) return false;
                if (element.Is(W, "tblStylePr")) return true;
                if (element.Is(W, "tblStyle") && WVal(element) is string tableStyleId && !string.IsNullOrWhiteSpace(tableStyleId)) {
                    usedTableStyleIds.Add(tableStyleId);
                }
                if (element.Is(W, "tbl") && ChildVal(element.Child(W, "tblPr"), "tblStyle") == null) hasTableWithoutStyle = true;
                return false;
            });
            if (documentScan != ShapeScanResult.Absent) return documentScan;

            IReadOnlyList<ComparisonStyle> styles = CollectComparisonStyles(mainPart, comparisonWorkBudget, out bool styleBudgetExhausted);
            if (styleBudgetExhausted) return ShapeScanResult.ResourceLimitExceeded;
            if (hasTableWithoutStyle) {
                AddDefaultStyleIds(styles, "table", usedTableStyleIds);
            }
            if (usedTableStyleIds.Count == 0) return ShapeScanResult.Absent;

            ILookup<string, ComparisonStyle> stylesById = styles
                .Where(style => !string.IsNullOrWhiteSpace(style.StyleId))
                .ToLookup(style => style.StyleId!, StringComparer.OrdinalIgnoreCase);
            return ContainsStyleShape(
                usedTableStyleIds,
                stylesById,
                (style, budget) => ScanElementSubtree(style, budget, element => element.Is(W, "tblStylePr") && !IsInsideExtensionList(element)),
                comparisonWorkBudget);
        }

        private static ShapeScanResult ContainsStyleShape(
            IEnumerable<string> initialStyleIds,
            ILookup<string, ComparisonStyle> stylesById,
            Func<OfficeXmlNode, ComparisonWorkBudget, ShapeScanResult> containsShape,
            ComparisonWorkBudget comparisonWorkBudget) {
            var pendingStyleIds = new Stack<string>(initialStyleIds);
            var inspectedStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (pendingStyleIds.Count > 0) {
                string styleId = pendingStyleIds.Pop();
                if (!inspectedStyleIds.Add(styleId)) continue;

                foreach (ComparisonStyle style in stylesById[styleId]) {
                    if (!comparisonWorkBudget.TryConsume(1)) return ShapeScanResult.ResourceLimitExceeded;
                    ShapeScanResult styleScan = containsShape(style.Element, comparisonWorkBudget);
                    if (styleScan != ShapeScanResult.Absent) return styleScan;
                    string? baseStyleId = style.BasedOn;
                    if (!string.IsNullOrWhiteSpace(baseStyleId)) pendingStyleIds.Push(baseStyleId!);
                }
            }
            return ShapeScanResult.Absent;
        }

        private static ShapeScanResult ContainsNumberingFormatting(
            WordPartView mainPart,
            ComparisonWorkBudget comparisonWorkBudget,
            ParagraphNumberingStyleCatalogCache numberingStyleCatalogs) {
            if (!numberingStyleCatalogs.TryGetOrCreate(
                    mainPart,
                    comparisonWorkBudget,
                    out ParagraphNumberingStyleCatalog? styleCatalog)) {
                return ShapeScanResult.ResourceLimitExceeded;
            }
            return ScanComparisonElements(mainPart, comparisonWorkBudget, element => {
                if (!element.Is(W, "p") || IsInsideExtensionList(element)) return ShapeScanResult.Absent;
                if (!TryResolveParagraphNumberingProperties(
                        element,
                        styleCatalog!,
                        comparisonWorkBudget,
                        out OfficeXmlNode numbering)) {
                    return ShapeScanResult.ResourceLimitExceeded;
                }
                return numbering.IsNone ? ShapeScanResult.Absent : ShapeScanResult.Present;
            });
        }

        private static ShapeScanResult ContainsMoveMarkup(WordPartView mainPart, ComparisonWorkBudget comparisonWorkBudget) =>
            ScanComparisonElements(mainPart, comparisonWorkBudget, element =>
                element.LocalName.StartsWith("moveFrom", StringComparison.Ordinal) ||
                element.LocalName.StartsWith("moveTo", StringComparison.Ordinal));

        private static ShapeScanResult ScanComparisonElements(
            WordPartView mainPart,
            ComparisonWorkBudget comparisonWorkBudget,
            Func<OfficeXmlNode, bool> predicate) => ScanComparisonElements(
                mainPart,
                comparisonWorkBudget,
                element => predicate(element) ? ShapeScanResult.Present : ShapeScanResult.Absent);

        private static ShapeScanResult ScanComparisonElements(
            WordPartView mainPart,
            ComparisonWorkBudget comparisonWorkBudget,
            Func<OfficeXmlNode, ShapeScanResult> inspect) {
            foreach (OfficeXmlNode root in EnumerateComparisonRoots(mainPart)) {
                if (!comparisonWorkBudget.TryConsume(1)) return ShapeScanResult.ResourceLimitExceeded;
                ShapeScanResult rootResult = inspect(root);
                if (rootResult != ShapeScanResult.Absent) return rootResult;
                foreach (OfficeXmlNode element in Descendants(root)) {
                    if (!comparisonWorkBudget.TryConsume(1)) return ShapeScanResult.ResourceLimitExceeded;
                    ShapeScanResult elementResult = inspect(element);
                    if (elementResult != ShapeScanResult.Absent) return elementResult;
                }
            }
            return ShapeScanResult.Absent;
        }

        private static ShapeScanResult ScanElementSubtree(
            OfficeXmlNode root,
            ComparisonWorkBudget comparisonWorkBudget,
            Func<OfficeXmlNode, bool> predicate) {
            if (!comparisonWorkBudget.TryConsume(1)) return ShapeScanResult.ResourceLimitExceeded;
            if (predicate(root)) return ShapeScanResult.Present;
            foreach (OfficeXmlNode element in Descendants(root)) {
                if (!comparisonWorkBudget.TryConsume(1)) return ShapeScanResult.ResourceLimitExceeded;
                if (predicate(element)) return ShapeScanResult.Present;
            }
            return ShapeScanResult.Absent;
        }

        private static IReadOnlyList<ComparisonStyle> CollectComparisonStyles(
            WordPartView mainPart,
            ComparisonWorkBudget comparisonWorkBudget,
            out bool budgetExhausted) {
            var styles = new List<ComparisonStyle>();
            foreach (ComparisonStyle style in EnumerateComparisonStyles(mainPart)) {
                if (!comparisonWorkBudget.TryConsume(1)) {
                    budgetExhausted = true;
                    return styles;
                }
                styles.Add(style);
            }
            budgetExhausted = false;
            return styles;
        }

        /// <summary>The styles part's then the styles-with-effects part's styles.</summary>
        private static IEnumerable<ComparisonStyle> EnumerateComparisonStyles(WordPartView mainPart) =>
            GetComparisonStyles(mainPart).Definitions
                .Concat(GetComparisonStyles(mainPart, WordPackageView.StylesWithEffectsRelationshipType).Definitions);

        private static IEnumerable<OfficeXmlNode> EnumerateComparisonRoots(WordPartView mainPart) =>
            WordFieldInventory.EnumerateFieldRoots(mainPart).Select(root => root.Root);
    }
}
