using OfficeIMO.Drawing.Internal;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods for Word document manipulation.
    /// </summary>
    public static partial class WordHelpers {
        /// <summary>
        /// Converts a DOTX template to a DOCX document.
        ///
        /// Based on: https://github.com/onizet/html2openxml/wiki/Convert-.dotx-to-.docx
        /// </summary>
        /// <param name="templatePath">The path to the DOTX template file.</param>
        /// <param name="outputPath">The path where the converted DOCX file will be saved.</param>
        public static void ConvertDotxToDocx(string templatePath, string outputPath) {
            string fullTemplatePath = Path.GetFullPath(templatePath);
            byte[] templateBytes = File.ReadAllBytes(fullTemplatePath);
            using (var documentStream = new MemoryStream()) {
                documentStream.Write(templateBytes, 0, templateBytes.Length);
                documentStream.Position = 0;

                using (WordprocessingDocument template = WordprocessingDocument.Open(documentStream, true)) {
                    template.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

                    MainDocumentPart mainPart = template.MainDocumentPart ?? throw new InvalidOperationException("MainDocumentPart is missing in template.");
                    if (mainPart.DocumentSettingsPart == null) {
                        mainPart.AddNewPart<DocumentSettingsPart>();
                    }

                    mainPart.DocumentSettingsPart!.AddExternalRelationship(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate",
                        new Uri(fullTemplatePath, UriKind.Absolute));
                    mainPart.Document?.Save();
                }

                OfficeFileCommit.WriteAllBytes(outputPath, documentStream.ToArray());
            }
        }
    }
}
