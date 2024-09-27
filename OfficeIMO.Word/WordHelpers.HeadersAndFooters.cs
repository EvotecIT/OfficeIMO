using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public partial class WordHelpers {
        /// <summary>
        /// Converts a DOTX template to a DOCX document.
        ///
        /// Based on: https://github.com/onizet/html2openxml/wiki/Convert-.dotx-to-.docx
        /// </summary>
        /// <param name="templatePath">The path to the DOTX template file.</param>
        /// <param name="outputPath">The path where the converted DOCX file will be saved.</param>
        public static void ConvertDotXtoDocX(string templatePath, string outputPath) {
            // Read the DOTX template file into a MemoryStream
            MemoryStream documentStream = Helpers.ReadFileToMemoryStream(templatePath);

            // Open the WordprocessingDocument from the MemoryStream
            using (WordprocessingDocument template = WordprocessingDocument.Open(documentStream, true)) {
                // Change the document type from template to document
                template.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

                // Get the main document part
                MainDocumentPart mainPart = template.MainDocumentPart;

                // Ensure the DocumentSettingsPart exists
                if (mainPart.DocumentSettingsPart == null) {
                    mainPart.AddNewPart<DocumentSettingsPart>();
                }

                // Add an external relationship to the template
                mainPart.DocumentSettingsPart.AddExternalRelationship(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate",
                    new Uri(templatePath, UriKind.Absolute));

                // Save the changes to the main document part
                mainPart.Document.Save();
            }

            // Write the MemoryStream contents to the output file
            File.WriteAllBytes(outputPath, documentStream.ToArray());
        }
    }
}
