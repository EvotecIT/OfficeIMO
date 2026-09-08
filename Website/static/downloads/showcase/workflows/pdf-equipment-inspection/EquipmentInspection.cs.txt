using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a fillable equipment inspection form with named text, choice and checkbox fields.</summary>
internal static class EquipmentInspection {
    internal static void Create(string folder) {
        PdfColor navy = PdfColor.FromRgb(23, 54, 93);
        PdfDocument.Create(pdf => pdf.Content(content => content
            .H1("Equipment inspection", PdfAlign.Left, navy)
            .Paragraph(p => p.Text("NORTHWIND / WORKPLACE OPERATIONS / FILLABLE CHECKLIST"))
            .Paragraph(p => p.Text("Record the inspection before returning equipment to the shared pool. This sample is not a safety certification."))
            .H2("Asset and inspector", PdfAlign.Left, navy)
            .Paragraph(p => p.Bold("Asset identifier"))
            .TextField("asset-id", width: 280, value: "NW-LT-042")
            .Paragraph(p => p.Bold("Inspector"))
            .TextField("inspector", width: 280)
            .Paragraph(p => p.Bold("Inspection date (YYYY-MM-DD)"))
            .TextField("inspection-date", width: 180)
            .H2("Record the checks", PdfAlign.Left, navy)
            .Paragraph(p => p.Text("Casing and connectors checked"))
            .CheckBox("casing-checked")
            .Paragraph(p => p.Text("Power-on and basic operation checked"))
            .CheckBox("operation-checked")
            .Paragraph(p => p.Text("Accessories returned and recorded"))
            .CheckBox("accessories-checked")
            .Paragraph(p => p.Bold("Disposition"))
            .ChoiceField("disposition", new[] { "Awaiting inspection", "Return to pool", "Repair required", "Retire from use" },
                value: "Awaiting inspection", width: 280)
            .Paragraph(p => p.Bold("Follow-up note"))
            .TextField("follow-up", width: 450, height: 38)),
            new PdfOptions { DefaultFontSize = 10 })
            .Meta(title: "Equipment inspection", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
