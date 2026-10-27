using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates a fillable maintenance job card with named work, outcome and handover fields.</summary>
    internal static class MaintenanceJobCard {
        internal static void Create(string folder) {
            PdfColor muted = PdfColor.FromOfficeColor(OfficeTheme.Default.Palette.TextMuted);
            PdfDocument.Create(pdf => pdf.Content(content => {
                content.Title("Maintenance job card")
                    .Paragraph(p => p.FontSize(9).Text("Northwind · Workplace operations · Demonstration form"), defaultColor: muted)
                    .PanelParagraph(p => p.Bold("Scope").Text("\nRecord authorised work and its result. This example does not define a safety procedure or authorise equipment operation."))
                    .H2("Work order");
                EquipmentInspection.Field(content, "Reference and asset", c => c.TextField("job-reference", width: 300, fontSize: 11, value: "WO-018 / ROOM-04", style: EquipmentInspection.Named("Reference and asset")));
                EquipmentInspection.Field(content, "Assigned technician", c => c.TextField("technician", width: 300, fontSize: 11, style: EquipmentInspection.Named("Assigned technician")));
                EquipmentInspection.Field(content, "Planned work", c => c.TextField("planned-work", width: 450, height: 48, fontSize: 11,
                    value: "Inspect the meeting-room display and record the connection test.", style: EquipmentInspection.Named("Planned work", multiline: true)));
                content.H2("Record the result");
                EquipmentInspection.Field(content, "Work completed", c => c.TextField("completed-work", width: 450, height: 48, fontSize: 11, style: EquipmentInspection.Named("Work completed", multiline: true)));
                EquipmentInspection.Field(content, "Outcome", c => c.ChoiceField("outcome", new[] { "Awaiting work", "Completed", "Follow-up needed" }, value: "Awaiting work", width: 260, fontSize: 11, style: EquipmentInspection.Named("Outcome")));
                EquipmentInspection.Check(content, "handover-acknowledged", "Handover acknowledged");
                EquipmentInspection.Field(content, "Follow-up owner and date", c => c.TextField("follow-up", width: 450, fontSize: 11, style: EquipmentInspection.Named("Follow-up owner and date")));
            }), new PdfCreateOptions { ShowPageNumbers = true })
                .Meta(title: "Maintenance job card", author: "OfficeIMO").Save(Path.Combine(folder, "example.pdf"));
        }
    }
}
