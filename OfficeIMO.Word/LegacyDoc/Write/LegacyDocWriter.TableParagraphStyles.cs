namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static LegacyDocWritableParagraphFormatting ReadSupportedCellParagraphStyleFormatting(LegacyDocElement paragraph, LegacyDocSourcePackage source) {
            var formatting = LegacyDocWritableParagraphFormatting.Plain;
            foreach (LegacyDocElement style in EnumerateCellParagraphStyleChain(paragraph, source)) {
                string styleId = LegacyDocStyleCatalog.StyleId(style)!;
                LegacyDocWritableParagraphFormatting own = TryMapBuiltInParagraphStyleIndex(styleId, out ushort styleIndex)
                    ? ReadSupportedBuiltInStyleParagraphFormatting(styleIndex, style.Child("pPr"))
                    : ReadSupportedCustomParagraphStyleParagraphFormatting(style.Child("pPr"));
                formatting = formatting.WithInheritedParagraphFormatting(own);
            }
            return formatting;
        }

        private static LegacyDocWritableFormatting ReadSupportedCellParagraphStyleRunFormatting(LegacyDocElement paragraph, LegacyDocSourcePackage source) {
            var formatting = LegacyDocWritableFormatting.Plain;
            foreach (LegacyDocElement style in EnumerateCellParagraphStyleChain(paragraph, source)) {
                formatting = formatting.WithInheritedFormatting(ReadSupportedRunFormatting(GetSupportedTableStyleRunProperties(style)));
            }
            return formatting;
        }

        private static IEnumerable<LegacyDocElement> EnumerateCellParagraphStyleChain(LegacyDocElement paragraph, LegacyDocSourcePackage source) {
            string? styleId = paragraph.Child("pPr").Child("pStyle").Val;
            // Unstyled cells retain table defaults ahead of the document's Normal defaults.
            if (string.IsNullOrWhiteSpace(styleId) || string.Equals(styleId, "Normal", StringComparison.Ordinal)) {
                yield break;
            }

            var visited = new HashSet<string>(StringComparer.Ordinal);
            IReadOnlyList<LegacyDocElement> styles = source.Styles.Definitions;
            while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                LegacyDocElement style = LegacyDocElement.None;
                foreach (LegacyDocElement candidate in styles) {
                    if (LegacyDocStyleCatalog.IsParagraphStyle(candidate) &&
                        string.Equals(LegacyDocStyleCatalog.StyleId(candidate), styleId, StringComparison.Ordinal)) {
                        style = candidate;
                        break;
                    }
                }
                if (style.IsNone) break;
                yield return style;
                styleId = style.Child("basedOn").Val;
            }
        }
    }
}
