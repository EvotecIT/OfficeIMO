using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static class DocumentInspector {
        internal static void Example_PrepareForSharing(string folderPath, bool openWord) {
            Console.WriteLine("[*] Word - Remove personal information and check accessibility before sharing");
            string filePath = Path.Combine(folderPath, "Document Inspector.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Properties.Creator = "Jane Doe";
                document.ApplicationProperties.Company = "Contoso Legal";
                document.SetDocumentVariable("ClientMatter", "2026-0417");
                document.AddHeading("Settlement summary", 1);
                document.AddParagraph("Both parties accept the revised schedule.")
                    .AddComment("Jane Doe", "JD", "Confirm the dates with finance first.");
                document.AddTable(3, 2).Rows[0].Cells[0].Paragraphs[0].Text = "Milestone";
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                foreach (OfficeHiddenDataFinding finding in document.InspectHiddenData().Findings) {
                    Console.WriteLine($"    found {finding.Kind} at {finding.Location}: {finding.Description}");
                }

                OfficeHiddenDataReport removed = document.RemoveHiddenData(new OfficeHiddenDataOptions {
                    TrackedChanges = OfficeTrackedChangesAction.AcceptAll
                });
                Console.WriteLine($"    removed or anonymized {removed.Findings.Count(finding => finding.Outcome != OfficeHiddenDataOutcome.Retained)} items");

                OfficeAccessibilityReport accessibility = document.FixAccessibility(new WordAccessibilityFixOptions {
                    Title = "Settlement summary",
                    Language = "en-US",
                    MarkTableHeaderRows = true
                });
                foreach (OfficeAccessibilityFinding finding in accessibility.Findings) {
                    Console.WriteLine($"    {finding.Severity} {finding.Code} at {finding.Location}");
                }

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
