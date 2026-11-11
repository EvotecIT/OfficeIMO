using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>Removes speaker notes and personal metadata, and anonymizes kept comments, before sharing a deck.</summary>
    public static class DocumentInspector {
        public static void Example_PrepareForSharing(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Remove speaker notes and anonymize comments before sharing");
            string filePath = Path.Combine(folderPath, "Document Inspector.pptx");
            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTextBox("Pricing proposal", OfficeLength.Points(40), OfficeLength.Points(40), OfficeLength.Points(400), OfficeLength.Points(50));
                slide.Notes.Text = "Our floor price is 20% lower.";
                presentation.AddModernComment(slide, new PowerPointCommentAuthor("Jane Doe", "JD"), "Approved by sales.");
                presentation.Properties.Creator = "Jane Doe";
                presentation.Save();
            }

            using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
                OfficeHiddenDataReport report = presentation.RemoveHiddenData(new OfficeHiddenDataOptions {
                    Kinds = OfficeHiddenDataKind.Default & ~OfficeHiddenDataKind.Comments
                });
                foreach (OfficeHiddenDataFinding finding in report.Findings) {
                    Console.WriteLine($"    {finding.Outcome} {finding.Kind} at {finding.Location}: {finding.Description}");
                }
                presentation.Save();
                if (openPowerPoint) presentation.OpenInApplication();
            }
        }
    }
}
