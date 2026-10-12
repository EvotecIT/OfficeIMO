using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Establishes the binary DOC paragraph baseline before importing source styles and direct formatting.
        /// Omitted DOC spacing means single lines and no paragraph spacing, rather than the DOCX authoring defaults.
        /// </summary>
        private static void ApplyLegacyDocParagraphDefaults(WordDocument document) {
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            DocDefaults defaults = styles.DocDefaults ?? styles.PrependChild(new DocDefaults());
            ParagraphPropertiesDefault paragraphDefaults = defaults.ParagraphPropertiesDefault ??
                defaults.AppendChild(new ParagraphPropertiesDefault());
            ParagraphPropertiesBaseStyle properties = paragraphDefaults.ParagraphPropertiesBaseStyle ??
                paragraphDefaults.AppendChild(new ParagraphPropertiesBaseStyle());
            properties.RemoveAllChildren<SpacingBetweenLines>();
            properties.AddChild(new SpacingBetweenLines {
                Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto
            }, true);
        }
    }
}
