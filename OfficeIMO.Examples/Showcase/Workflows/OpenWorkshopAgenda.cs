using OfficeIMO.OpenDocument;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Saves the same workshop agenda as packaged ODT and flat XML FODT for inspection and exchange.</summary>
internal static class OpenWorkshopAgenda {
    internal static void Create(string folder) {
        OdtDocument document = OdtDocument.Create();
        var title = document.AddHeading("Service design workshop", 1);
        title.FontFamily = "Carlito"; title.FontSize = OdfLength.Parse("26pt"); title.Color = OdfColor.Parse("#17365d");
        document.AddParagraph("90 minutes / six participants / facilitator: Jordan Lee");
        document.AddHeading("Outcome", 2);
        document.AddParagraph("Agree one request journey, identify the unresolved decisions and leave with named owners for the next actions.");
        document.AddHeading("Agenda", 2);
        Activity("00–10", "Frame the problem", "Agree the outcome and boundaries.");
        Activity("10–30", "Map the current journey", "Make handoffs and waiting points visible.");
        Activity("30–55", "Design the next journey", "Name roles and decision points.");
        Activity("55–75", "Test two scenarios", "Record exceptions and open questions.");
        Activity("75–90", "Commit next actions", "Agree owners, dates and a review point.");
        document.AddHeading("Preparation", 2);
        document.AddParagraph("Bring two anonymized requests: one routine case and one that needed several handoffs. Read the current service description before the session.");
        document.AddHeading("Capture before closing", 2);
        var close = document.AddParagraph();
        close.AddSpan("Decision / owner / due date").Bold = true;
        close.AddSpan(" — record all three for every action. Share the agenda and the resulting notes with participants.");
        foreach (var paragraph in document.Paragraphs) {
            paragraph.FontFamily = "Carlito";
            paragraph.FontSize = OdfLength.Parse(paragraph.HeadingLevel == 1 ? "26pt" : paragraph.IsHeading ? "15pt" : "11pt");
            paragraph.SpaceAbove = OdfLength.Parse(paragraph.IsHeading ? "12pt" : "0pt");
            paragraph.SpaceBelow = OdfLength.Parse("8pt");
            if (paragraph.IsHeading) paragraph.Color = OdfColor.Parse("#17365d");
        }
        document.Save(Path.Combine(folder, "example.odt"));
        document.SaveFlatXml(Path.Combine(folder, "example.fodt"));

        void Activity(string time, string title, string outcome) {
            var paragraph = document.AddParagraph();
            var label = paragraph.AddSpan(time + "  " + title + ". ");
            label.Bold = true;
            label.Color = OdfColor.Parse("#2563eb");
            paragraph.AddSpan(outcome);
        }
    }
}
