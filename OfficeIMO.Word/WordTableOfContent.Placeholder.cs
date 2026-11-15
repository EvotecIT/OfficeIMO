using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTableOfContent {
        private string GetPlaceholderText() {
            DocPart? placeholder = GetPlaceholderPart();
            if (placeholder != null) return string.Concat(placeholder.DocPartBody?.Descendants<Text>().Select(text => text.Text) ?? Enumerable.Empty<string>());
            return _sdtBlock.SdtContentBlock?.Elements<Paragraph>()
                .SelectMany(paragraph => paragraph.Elements<SimpleField>())
                .FirstOrDefault(field => IsGeneratedListInstruction(field.Instruction?.Value))?
                .Descendants<Text>().FirstOrDefault()?.Text ?? string.Empty;
        }

        private DocPart? GetPlaceholderPart() {
            string? name = _sdtBlock.SdtProperties?.GetFirstChild<SdtPlaceholder>()?.DocPartReference?.Val?.Value;
            return name == null ? null : _document._wordprocessingDocument.MainDocumentPart?.GlossaryDocumentPart?
                .GlossaryDocument?.DocParts?.Elements<DocPart>().FirstOrDefault(part => part.DocPartProperties?.DocPartName?.Val?.Value == name);
        }

        private void SetPlaceholderText(string value) {
            PreservePlaceholderText(value);
            SdtContentBlock? content = _sdtBlock.SdtContentBlock;
            if (content == null || content.Elements<Paragraph>().Any(IsGeneratedResultParagraph)) return;
            SimpleField? simple = content.Descendants<SimpleField>().FirstOrDefault(field => IsGeneratedListInstruction(field.Instruction?.Value));
            Text? text = simple?.Descendants<Text>().FirstOrDefault();
            if (text == null) {
                text = content.Elements<Paragraph>().Where(paragraph => paragraph.Descendants<FieldCode>().Any(code => IsGeneratedListInstruction(code.Text)))
                    .SelectMany(paragraph => paragraph.Descendants<Text>()).FirstOrDefault();
            }
            if (text != null) text.Text = value;
        }

        private static bool IsGeneratedListInstruction(string? instruction) {
            WordFieldType? type = string.IsNullOrWhiteSpace(instruction) ? null : WordFieldInventory.ParseInstruction(instruction!).FieldType;
            return type == WordFieldType.TOC || type == WordFieldType.Index;
        }

        private void PreservePlaceholderText(string value) {
            // Use Word's glossary placeholder so the preference survives replacement of the field result.
            DocPart? part = GetPlaceholderPart();
            if (part == null || !(part.DocPartProperties?.DocPartName?.Val?.Value?.StartsWith("OfficeIMO_TocPlaceholder_", StringComparison.Ordinal) ?? false)) {
                MainDocumentPart main = _document._wordprocessingDocument.MainDocumentPart!;
                GlossaryDocumentPart glossary = main.GlossaryDocumentPart ?? main.AddNewPart<GlossaryDocumentPart>();
                glossary.GlossaryDocument ??= new GlossaryDocument();
                glossary.GlossaryDocument.DocParts ??= new DocParts();
                string name = "OfficeIMO_TocPlaceholder_" + Guid.NewGuid().ToString("N");
                part = new DocPart(new DocPartProperties(new DocPartName { Val = name },
                    new DocPartTypes(new DocPartType { Val = DocPartValues.SdtPlaceholder })));
                glossary.GlossaryDocument.DocParts.Append(part);
                _sdtBlock.SdtProperties ??= new SdtProperties();
                _sdtBlock.SdtProperties.RemoveAllChildren<SdtPlaceholder>();
                _sdtBlock.SdtProperties.Append(new SdtPlaceholder(new DocPartReference { Val = name }));
            }
            part.DocPartBody = new DocPartBody(new Paragraph(new Run(new Text(value) { Space = SpaceProcessingModeValues.Preserve })));
        }

        internal void CopyGeneratedRefreshStateTo(WordTableOfContent clone) {
            clone._entriesWrittenOnSave = _entriesWrittenOnSave;
        }
    }
}
