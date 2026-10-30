using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates an inspection form, fills a completed copy, and saves a permanent record.</summary>
internal static class EquipmentInspection {
    internal static void Create(string folder) {
        global::OfficeIMO.OfficeColor navy = global::OfficeIMO.OfficeColor.FromRgb(23, 54, 93);
        PdfDocument template = PdfDocument.Create(pdf => pdf.Content(content => content
            .H1("Equipment inspection", PdfAlign.Left, navy)
            .Paragraph(p => p.Text("NORTHWIND / WORKPLACE OPERATIONS / FILLABLE CHECKLIST"))
            .Paragraph(p => p.Text("Record the inspection before returning equipment to the shared pool. This sample is not a safety certification."))
            .H2("Asset and inspector", PdfAlign.Left, navy)
            .Paragraph(p => p.Bold("Asset identifier"))
            .TextField("asset-id", width: global::OfficeIMO.OfficeLength.Points(280), value: "NW-LT-042")
            .Paragraph(p => p.Bold("Inspector"))
            .TextField("inspector", width: global::OfficeIMO.OfficeLength.Points(280))
            .Paragraph(p => p.Bold("Inspection date (YYYY-MM-DD)"))
            .TextField("inspection-date", width: global::OfficeIMO.OfficeLength.Points(180))
            .H2("Record the checks", PdfAlign.Left, navy)
            .Paragraph(p => p.Text("Casing and connectors checked"))
            .CheckBox("casing-checked")
            .Paragraph(p => p.Text("Power-on and basic operation checked"))
            .CheckBox("operation-checked")
            .Paragraph(p => p.Text("Accessories returned and recorded"))
            .CheckBox("accessories-checked")
            .Paragraph(p => p.Bold("Disposition"))
            .ChoiceField("disposition", new[] { "Awaiting inspection", "Return to pool", "Repair required", "Retire from use" },
                value: "Awaiting inspection", width: global::OfficeIMO.OfficeLength.Points(280))
            .Paragraph(p => p.Bold("Follow-up note"))
            .TextField("follow-up", width: global::OfficeIMO.OfficeLength.Points(450), height: global::OfficeIMO.OfficeLength.Points(38))),
            new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) })
            .Meta(title: "Equipment inspection", author: "OfficeIMO");
        template.Save(Path.Combine(folder, "example.pdf"));

        PdfDocument completed = template.Forms.Fill(new Dictionary<string, string> {
            ["inspector"] = "Marta Nowak",
            ["inspection-date"] = "2026-09-17",
            ["casing-checked"] = "Yes",
            ["operation-checked"] = "Yes",
            ["accessories-checked"] = "Yes",
            ["disposition"] = "Return to pool",
            ["follow-up"] = "All accessories returned. Ready for the next booking."
        });
        completed.Save(Path.Combine(folder, "completed.pdf"));
        completed.Forms.Flatten().Save(Path.Combine(folder, "record.pdf"));
    }
}
