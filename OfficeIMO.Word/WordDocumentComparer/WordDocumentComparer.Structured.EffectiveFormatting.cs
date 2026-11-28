using System.Text;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static string GetParagraphFormatSignature(OfficeXmlNode paragraph, WordPartView? part, WordComparisonOptions options) {
            if (!options.CompareEffectiveFormatting) {
                return string.Empty;
            }

            ComparisonStyles styles = GetComparisonStyles(GetMainDocumentPart(part));
            var effectiveProperties = new List<OfficeXmlNode>();

            MergePropertyChildren(effectiveProperties, styles.DocumentDefaults.Child(W, "pPrDefault").Child(W, "pPr"), IncludeEffectiveParagraphProperty);
            MergeStyleProperties(effectiveProperties, styles, "paragraph", GetParagraphStyleId(paragraph), "pPr", IncludeEffectiveParagraphProperty, new HashSet<string>(StringComparer.Ordinal));
            MergePropertyChildren(effectiveProperties, paragraph.Child(W, "pPr"), IncludeEffectiveParagraphProperty);

            return GetPropertySetSignature("pPr", effectiveProperties);
        }

        private static string GetEffectiveRunFormatSignature(OfficeXmlNode run, OfficeXmlNode paragraph, WordPartView? part, WordComparisonOptions options) {
            ComparisonStyles styles = GetComparisonStyles(GetMainDocumentPart(part));
            var effectiveProperties = new List<OfficeXmlNode>();

            MergePropertyChildren(effectiveProperties, styles.DocumentDefaults.Child(W, "rPrDefault").Child(W, "rPr"), IncludeEffectiveRunProperty);
            MergeStyleProperties(effectiveProperties, styles, "paragraph", GetParagraphStyleId(paragraph), "rPr", IncludeEffectiveRunProperty, new HashSet<string>(StringComparer.Ordinal));

            OfficeXmlNode runProperties = run.Child(W, "rPr");
            string runStyleId = ChildVal(runProperties, "rStyle") ?? string.Empty;
            MergeStyleProperties(effectiveProperties, styles, "character", runStyleId, "rPr", IncludeEffectiveRunProperty, new HashSet<string>(StringComparer.Ordinal));
            MergePropertyChildren(effectiveProperties, runProperties, element => IncludeEffectiveRunProperty(element, options));

            return GetPropertySetSignature("rPr", effectiveProperties);
        }

        private static string GetPropertySetSignature(string container, List<OfficeXmlNode> properties) {
            if (properties.Count == 0) {
                return string.Empty;
            }

            var builder = new StringBuilder(container).Append('>');
            foreach (OfficeXmlNode property in properties) {
                builder.Append(GetElementSignature(property));
            }

            return builder.ToString();
        }

        private static void MergeStyleProperties(
            List<OfficeXmlNode> target,
            ComparisonStyles styles,
            string styleType,
            string styleId,
            string propertiesName,
            Func<OfficeXmlNode, bool> include,
            ISet<string> visitedStyleIds) {
            ComparisonStyle? style = FindStyle(styles, styleType, styleId);
            if (style == null || !visitedStyleIds.Add(style.StyleId ?? string.Empty)) {
                return;
            }

            string basedOnStyleId = style.BasedOn ?? string.Empty;
            if (basedOnStyleId.Length > 0) {
                MergeStyleProperties(target, styles, styleType, basedOnStyleId, propertiesName, include, visitedStyleIds);
            }

            MergePropertyChildren(target, style.Element.Child(W, propertiesName), include);
        }

        // The first style of exactly this type with the identifier, otherwise the type's default style.
        private static ComparisonStyle? FindStyle(ComparisonStyles styles, string typeName, string styleId) {
            ComparisonStyle? explicitStyle = string.IsNullOrEmpty(styleId) ? null : styles.Definitions.FirstOrDefault(style =>
                string.Equals(style.Type, typeName, StringComparison.Ordinal) &&
                string.Equals(style.StyleId, styleId, StringComparison.Ordinal));
            return explicitStyle ?? styles.Definitions.FirstOrDefault(style =>
                string.Equals(style.Type, typeName, StringComparison.Ordinal) && style.IsDefault);
        }

        private static void MergePropertyChildren(List<OfficeXmlNode> target, OfficeXmlNode source, Func<OfficeXmlNode, bool> include) {
            if (source.IsNone) {
                return;
            }

            for (OfficeXmlNode child = source.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (!include(child)) {
                    continue;
                }

                string localName = child.LocalName;
                string namespaceUri = child.NamespaceUri;
                target.RemoveAll(existing =>
                    string.Equals(existing.LocalName, localName, StringComparison.Ordinal) &&
                    string.Equals(existing.NamespaceUri, namespaceUri, StringComparison.Ordinal));
                target.Add(child);
            }
        }

        private static bool IncludeEffectiveParagraphProperty(OfficeXmlNode element) {
            return !element.Is(W, "pStyle") && !element.Is(W, "pPrChange");
        }

        private static bool IncludeEffectiveRunProperty(OfficeXmlNode element) {
            return !element.Is(W, "rPrChange");
        }

        private static bool IncludeEffectiveRunProperty(OfficeXmlNode element, WordComparisonOptions options) {
            if (element.Is(W, "rStyle") && !options.CompareRunStyleIds) {
                return false;
            }

            return IncludeEffectiveRunProperty(element);
        }

        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<WordPartView, ComparisonStyles> StylesByPart = new();

        /// <summary>The styles of the package's styles part, read once per package view.</summary>
        private static ComparisonStyles GetComparisonStyles(WordPartView? mainPart) => GetComparisonStyles(mainPart, WordPackageView.StylesRelationshipType);

        private static ComparisonStyles GetComparisonStyles(WordPartView? mainPart, string relationshipType) {
            WordPartView? part = mainPart?.FirstPart(relationshipType);
            return part == null ? ComparisonStyles.Empty : StylesByPart.GetValue(part, stylesPart => new ComparisonStyles(stylesPart.RootNamed(W, "styles")));
        }
    }
}
