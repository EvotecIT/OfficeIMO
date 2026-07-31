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

        private static bool ContainsThemeFormatting(MainDocumentPart mainPart) =>
            EnumerateComparisonRoots(mainPart)
                .SelectMany(root => new[] { root }.Concat(root.Descendants()))
                .SelectMany(element => element.GetAttributes())
                .Any(attribute => attribute.LocalName.EndsWith("Theme", StringComparison.OrdinalIgnoreCase) ||
                                  attribute.LocalName.Equals("themeColor", StringComparison.OrdinalIgnoreCase));

        private static bool ContainsConditionalTableStyleFormatting(MainDocumentPart mainPart) =>
            EnumerateComparisonRoots(mainPart).Any(root =>
                root.Descendants<TableStyle>().Any() ||
                root.Descendants<TableStyleConditionalFormattingTableRowProperties>().Any() ||
                root.Descendants<TableStyleConditionalFormattingTableCellProperties>().Any());

        private static bool ContainsNumberingFormatting(MainDocumentPart mainPart) =>
            EnumerateComparisonRoots(mainPart).Any(root => root.Descendants<NumberingProperties>().Any());

        private static bool ContainsMoveMarkup(MainDocumentPart mainPart) =>
            EnumerateComparisonRoots(mainPart)
                .SelectMany(root => root.Descendants())
                .Any(element => element.LocalName.StartsWith("moveFrom", StringComparison.Ordinal) ||
                                element.LocalName.StartsWith("moveTo", StringComparison.Ordinal));

        private static IEnumerable<OpenXmlCompositeElement> EnumerateComparisonRoots(MainDocumentPart mainPart) =>
            WordFieldInventory.EnumerateFieldRoots(mainPart).Select(root => root.Root);
    }
}
