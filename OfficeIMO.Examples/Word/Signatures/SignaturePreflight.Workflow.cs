using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static class SignaturePreflight {
        internal static void Example_SignaturePreflightWorkflow(string folderPath, bool openWord) {
            Console.WriteLine("[*] Word signature preflight workflow");

            string documentPath = Path.Combine(folderPath, "SignaturePreflightWorkflow.docx");
            using (WordDocument document = WordDocument.Create(documentPath)) {
                document.AddParagraph("Signed package metadata preflight").Style = WordParagraphStyles.Heading1;
                document.AddParagraph("OfficeIMO can inspect signature package metadata and block accidental saves by default.");
                document.Save();
                if (openWord) document.OpenInApplication();
            }

            PremiumWorkflowExampleUtilities.AddSyntheticSignatureMetadata(documentPath);

            WordSignatureValidationReport validationReport;
            using (WordDocument document = WordDocument.Load(documentPath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                validationReport = document.ValidateSignatures();
            }

            string savePolicyMessage;
            using (WordDocument document = WordDocument.Load(documentPath)) {
                document.AddParagraph("This edit is intentionally blocked by the default signed-document save policy.");
                try {
                    document.Save();
                    savePolicyMessage = "Save unexpectedly succeeded.";
                } catch (WordSignatureSavePolicyException ex) {
                    savePolicyMessage = ex.Message;
                }
            }

            PremiumWorkflowExampleUtilities.WriteSignaturePreflightReport(
                Path.Combine(folderPath, "SignaturePreflightWorkflow.md"),
                validationReport,
                savePolicyMessage);
        }
    }
}
