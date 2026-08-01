using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructureReportsThemeLimitationFromDuplicateEffectsStyleDefinition() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_effects_style_theme_source.docx");
            CreateDocumentWithInheritedComparisonStyle(sourcePath, paragraphSpacingAfter: "120", runColor: "1F4E79");
            RemovePrimaryStyleThemeAttributes(sourcePath);
            AddEffectsStyleThemeColor(sourcePath, "OfficeIMOEffectiveBase", ThemeColorValues.Accent1);

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_effects_style_theme_target.docx");
            CreateDocumentWithInheritedComparisonStyle(targetPath, paragraphSpacingAfter: "120", runColor: "1F4E79");
            RemovePrimaryStyleThemeAttributes(targetPath);
            AddEffectsStyleThemeColor(targetPath, "OfficeIMOEffectiveBase", ThemeColorValues.Accent2);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Limitations, limitation =>
                limitation.Code == "EffectiveFormatting.ThemeResolution" &&
                limitation.SourceContainsShape &&
                limitation.TargetContainsShape);
        }

        [Fact]
        public void CompareStructureReportsConditionalTableLimitFromDuplicateEffectsStyleDefinition() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_effects_conditional_table_source.docx");
            CreateDocumentWithComparisonTableStyle(sourcePath, includeConditionalBaseStyle: false);
            AddEffectsConditionalTableStyle(sourcePath, "OfficeIMOTableBase");

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_effects_conditional_table_target.docx");
            CreateDocumentWithComparisonTableStyle(targetPath, includeConditionalBaseStyle: false);
            AddEffectsConditionalTableStyle(targetPath, "OfficeIMOTableBase");

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Limitations, limitation =>
                limitation.Code == "EffectiveFormatting.ConditionalTableStyles" &&
                limitation.SourceContainsShape &&
                limitation.TargetContainsShape);
        }

        private static void AddEffectsStyleThemeColor(string path, string styleId, ThemeColorValues themeColor) {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
            StylesWithEffectsPart effectsPart = mainPart.StylesWithEffectsPart ?? mainPart.AddNewPart<StylesWithEffectsPart>();
            effectsPart.Styles ??= new Styles();
            effectsPart.Styles.Append(new Style(
                new StyleName { Val = "OfficeIMO Effects Duplicate" },
                new StyleRunProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.Color { ThemeColor = themeColor })) {
                Type = StyleValues.Paragraph,
                StyleId = styleId,
                CustomStyle = true
            });
            effectsPart.Styles.Save();
        }

        private static void RemovePrimaryStyleThemeAttributes(string path) {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            Styles styles = wordDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            foreach (OpenXmlElement element in new[] { styles }.Concat(styles.Descendants())) {
                foreach (OpenXmlAttribute attribute in element.GetAttributes()
                    .Where(attribute =>
                        attribute.LocalName.EndsWith("Theme", System.StringComparison.OrdinalIgnoreCase) ||
                        attribute.LocalName.Equals("themeColor", System.StringComparison.OrdinalIgnoreCase))
                    .ToArray()) {
                    element.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
                }
            }
            styles.Save();
        }

        private static void AddEffectsConditionalTableStyle(string path, string styleId) {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
            StylesWithEffectsPart effectsPart = mainPart.StylesWithEffectsPart ?? mainPart.AddNewPart<StylesWithEffectsPart>();
            effectsPart.Styles ??= new Styles();
            effectsPart.Styles.Append(new Style(
                new StyleName { Val = "OfficeIMO Effects Conditional Duplicate" },
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading { Fill = "D9EAF7" })) {
                    Type = TableStyleOverrideValues.FirstRow
                }) {
                Type = StyleValues.Table,
                StyleId = styleId,
                CustomStyle = true
            });
            effectsPart.Styles.Save();
        }
    }
}
