using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static string GetParagraphFormatSignature(Paragraph paragraph, OpenXmlPart? part, WordComparisonOptions options) {
            if (!options.CompareEffectiveFormatting) {
                return string.Empty;
            }

            MainDocumentPart? mainPart = GetMainDocumentPart(part);
            Styles? styles = mainPart?.StyleDefinitionsPart?.Styles;
            var effectiveProperties = new ParagraphProperties();

            MergePropertyChildren(effectiveProperties, styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle, IncludeEffectiveParagraphProperty);
            MergeParagraphStyleProperties(effectiveProperties, styles, GetParagraphStyleId(paragraph), new HashSet<string>(StringComparer.Ordinal));
            MergePropertyChildren(effectiveProperties, paragraph.ParagraphProperties, IncludeEffectiveParagraphProperty);

            return effectiveProperties.HasChildren ? effectiveProperties.OuterXml : string.Empty;
        }

        private static string GetEffectiveRunFormatSignature(Run run, Paragraph paragraph, OpenXmlPart? part, WordComparisonOptions options) {
            MainDocumentPart? mainPart = GetMainDocumentPart(part);
            Styles? styles = mainPart?.StyleDefinitionsPart?.Styles;
            var effectiveProperties = new RunProperties();

            MergePropertyChildren(effectiveProperties, styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle, IncludeEffectiveRunProperty);
            MergeParagraphStyleRunProperties(effectiveProperties, styles, GetParagraphStyleId(paragraph), new HashSet<string>(StringComparer.Ordinal));

            string runStyleId = run.RunProperties?.RunStyle?.Val?.Value ?? string.Empty;
            MergeCharacterStyleRunProperties(effectiveProperties, styles, runStyleId, new HashSet<string>(StringComparer.Ordinal));
            MergePropertyChildren(effectiveProperties, run.RunProperties, element => IncludeEffectiveRunProperty(element, options));

            return effectiveProperties.HasChildren ? effectiveProperties.OuterXml : string.Empty;
        }

        private static void MergeParagraphStyleProperties(ParagraphProperties target, Styles? styles, string styleId, ISet<string> visitedStyleIds) {
            Style? style = FindStyle(styles, StyleValues.Paragraph, styleId);
            if (style == null || !visitedStyleIds.Add(style.StyleId?.Value ?? string.Empty)) {
                return;
            }

            string basedOnStyleId = style.BasedOn?.Val?.Value ?? string.Empty;
            if (basedOnStyleId.Length > 0) {
                MergeParagraphStyleProperties(target, styles, basedOnStyleId, visitedStyleIds);
            }

            MergePropertyChildren(target, style.StyleParagraphProperties, IncludeEffectiveParagraphProperty);
        }

        private static void MergeParagraphStyleRunProperties(RunProperties target, Styles? styles, string styleId, ISet<string> visitedStyleIds) {
            Style? style = FindStyle(styles, StyleValues.Paragraph, styleId);
            if (style == null || !visitedStyleIds.Add(style.StyleId?.Value ?? string.Empty)) {
                return;
            }

            string basedOnStyleId = style.BasedOn?.Val?.Value ?? string.Empty;
            if (basedOnStyleId.Length > 0) {
                MergeParagraphStyleRunProperties(target, styles, basedOnStyleId, visitedStyleIds);
            }

            MergePropertyChildren(target, style.StyleRunProperties, IncludeEffectiveRunProperty);
        }

        private static void MergeCharacterStyleRunProperties(RunProperties target, Styles? styles, string styleId, ISet<string> visitedStyleIds) {
            Style? style = FindStyle(styles, StyleValues.Character, styleId);
            if (style == null || !visitedStyleIds.Add(style.StyleId?.Value ?? string.Empty)) {
                return;
            }

            string basedOnStyleId = style.BasedOn?.Val?.Value ?? string.Empty;
            if (basedOnStyleId.Length > 0) {
                MergeCharacterStyleRunProperties(target, styles, basedOnStyleId, visitedStyleIds);
            }

            MergePropertyChildren(target, style.StyleRunProperties, IncludeEffectiveRunProperty);
        }

        private static Style? FindStyle(Styles? styles, StyleValues type, string styleId) {
            if (styles == null) {
                return null;
            }

            if (!string.IsNullOrEmpty(styleId)) {
                Style? explicitStyle = styles.Elements<Style>()
                    .FirstOrDefault(style =>
                        style.Type?.Value == type &&
                        string.Equals(style.StyleId?.Value, styleId, StringComparison.Ordinal));
                if (explicitStyle != null) {
                    return explicitStyle;
                }
            }

            return styles.Elements<Style>()
                .FirstOrDefault(style =>
                    style.Type?.Value == type &&
                    style.Default?.Value == true);
        }

        private static void MergePropertyChildren(OpenXmlCompositeElement target, OpenXmlCompositeElement? source, Func<OpenXmlElement, bool> include) {
            if (source == null) {
                return;
            }

            foreach (OpenXmlElement child in source.ChildElements) {
                if (!include(child)) {
                    continue;
                }

                OpenXmlElement clone = child.CloneNode(true);
                RemoveExistingProperty(target, clone);
                target.Append(clone);
            }
        }

        private static void RemoveExistingProperty(OpenXmlCompositeElement target, OpenXmlElement replacement) {
            foreach (OpenXmlElement existing in target.ChildElements
                .Where(child =>
                    string.Equals(child.LocalName, replacement.LocalName, StringComparison.Ordinal) &&
                    string.Equals(child.NamespaceUri, replacement.NamespaceUri, StringComparison.Ordinal))
                .ToList()) {
                existing.Remove();
            }
        }

        private static bool IncludeEffectiveParagraphProperty(OpenXmlElement element) {
            return element is not ParagraphStyleId &&
                   element is not ParagraphPropertiesChange;
        }

        private static bool IncludeEffectiveRunProperty(OpenXmlElement element) {
            return element is not RunPropertiesChange;
        }

        private static bool IncludeEffectiveRunProperty(OpenXmlElement element, WordComparisonOptions options) {
            if (element is RunStyle && !options.CompareRunStyleIds) {
                return false;
            }

            return IncludeEffectiveRunProperty(element);
        }
    }
}
