using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates an inspection form, fills a completed copy, and saves a permanent record.</summary>
internal static class EquipmentInspection {
    internal static void Create(string folder) {
        OfficeColor muted = OfficeTheme.Default.Palette.TextMuted;
        PdfDocument template = PdfDocument.Create(pdf => pdf.Content(content => {
            content.Title("Equipment inspection")
                .Paragraph(p => p.FontSize(OfficeLength.Points(9)).Text("Northwind · Workplace operations · Fillable checklist"), defaultColor: muted)
                .Paragraph(p => p.Text("Record the inspection before returning equipment to the shared pool. This sample is not a safety certification."))
                .H2("Asset and inspector");
            Field(content, "Asset identifier", c => c.TextField("asset-id", width: OfficeLength.Points(280), value: "NW-LT-042", style: Named("Asset identifier")));
            Field(content, "Inspector", c => c.TextField("inspector", width: OfficeLength.Points(280), style: Named("Inspector")));
            Field(content, "Inspection date (YYYY-MM-DD)", c => c.TextField("inspection-date", width: OfficeLength.Points(180), style: Named("Inspection date")));
            content.H2("Record the checks");
            Check(content, "casing-checked", "Casing and connectors checked");
            Check(content, "operation-checked", "Power-on and basic operation checked");
            Check(content, "accessories-checked", "Accessories returned and recorded");
            Field(content, "Disposition", c => c.ChoiceField("disposition", new[] { "Awaiting inspection", "Return to pool", "Repair required", "Retire from use" },
                value: "Awaiting inspection", width: OfficeLength.Points(280), style: Named("Disposition")));
            Field(content, "Follow-up note", c => c.TextField("follow-up", width: OfficeLength.Points(450), height: OfficeLength.Points(48), style: Named("Follow-up note", multiline: true)));
        }), new PdfCreateOptions { DefaultFontSize = OfficeLength.Points(10), ShowPageNumbers = true })
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

    // A label sits close to its own field and well clear of the field above it.
    internal static void Field(PdfContentBuilder content, string label, System.Action<PdfContentBuilder> field) {
        content.Paragraph(p => p.Bold(label), style: new PdfParagraphStyle { SpacingBefore = OfficeLength.Points(10), SpacingAfter = OfficeLength.Points(3), KeepWithNext = true });
        field(content);
    }

    internal static void Check(PdfContentBuilder content, string name, string label) =>
        content.Row(row => row.Style(new PdfRowStyle { SpacingBefore = OfficeLength.Points(6), SpacingAfter = OfficeLength.Points(2) }).Gap(OfficeLength.Points(8))
            .FixedColumn(OfficeLength.Points(16), column => column.CheckBox(name, spacingAfter: OfficeLength.Points(0), style: Named(label)))
            .RelativeColumn(column => column.Paragraph(p => p.Text(label), style: new PdfParagraphStyle { SpacingBefore = OfficeLength.Points(3), SpacingAfter = OfficeLength.Points(0) })));

    internal static PdfFormFieldStyle Named(string label, bool multiline = false) =>
        new PdfFormFieldStyle { AlternateName = label, IsMultiline = multiline };
}
