using OfficeIMO.OpenDocument;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Authors a vendor-neutral proposal with styled text, deliverables and explicit assumptions.</summary>
internal static class OpenProposal {
    internal static void Create(string folder) {
        OdtDocument document = OdtDocument.Create();
        var title = document.AddHeading("Service improvement proposal", 1);
        title.FontFamily = "Carlito"; title.FontSize = OdfLength.Parse("26pt"); title.Color = OdfColor.Parse("#17365d");
        document.AddParagraph("NORTHWIND / OPERATIONS / PROPOSAL 014");
        document.AddHeading("The proposed outcome", 2);
        document.AddParagraph("Reduce avoidable request handoffs by giving each service a clear intake path, named owner and agreed escalation route.");
        var decision = document.AddParagraph();
        var label = decision.AddSpan("Decision requested: "); label.Bold = true; label.Color = OdfColor.Parse("#2563eb");
        decision.AddSpan("approve a two-week discovery phase before committing to implementation.");
        document.AddHeading("Deliverables", 2);
        Deliverable("Current service map", "Operations", "Reviewed by the owners of each request queue.");
        Deliverable("Proposed intake journey", "Product", "Tested with five representative requests.");
        Deliverable("Implementation estimate", "Engineering", "Dependencies, assumptions and effort are stated.");
        document.AddHeading("Assumptions and exclusions", 2);
        document.AddParagraph("Queue owners are available for interviews. Discovery uses anonymized request examples. Procurement, platform replacement and production changes are outside this proposal.");
        document.AddHeading("Next step", 2);
        document.AddParagraph("Name the sponsor, agree the discovery dates and confirm who accepts each deliverable.");
        foreach (var paragraph in document.Paragraphs) {
            paragraph.FontFamily = "Carlito";
            paragraph.FontSize = OdfLength.Parse(paragraph.HeadingLevel == 1 ? "26pt" : paragraph.IsHeading ? "15pt" : "11pt");
            paragraph.SpaceAbove = OdfLength.Parse(paragraph.IsHeading ? "12pt" : "0pt");
            paragraph.SpaceBelow = OdfLength.Parse("7pt");
            if (paragraph.IsHeading) paragraph.Color = OdfColor.Parse("#17365d");
        }
        document.Save(Path.Combine(folder, "example.odt"));

        void Deliverable(string name, string owner, string acceptance) {
            var paragraph = document.AddParagraph();
            paragraph.AddSpan(name + " / " + owner + ". ").Bold = true;
            paragraph.AddSpan(acceptance);
        }
    }
}
