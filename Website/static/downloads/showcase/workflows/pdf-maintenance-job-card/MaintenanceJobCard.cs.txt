using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates a fillable maintenance job card with named work, outcome and handover fields.</summary>
    internal static class MaintenanceJobCard {
        internal static void Create(string folder) {
            var navy = global::OfficeIMO.OfficeColor.FromRgb(23, 54, 93);
            PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Maintenance job card", PdfAlign.Left, navy)
                .Paragraph(p => p.Text("NORTHWIND / WORKPLACE OPERATIONS / DEMONSTRATION FORM"))
                .PanelParagraph(p => p.Bold("Scope").Text("\nRecord authorised work and its result. This example does not define a safety procedure or authorise equipment operation."),
                    new PdfPanelStyle { Background = global::OfficeIMO.OfficeColor.FromRgb(234, 241, 251), PaddingX = 12, PaddingY = 10 })
                .H2("Work order", PdfAlign.Left, navy)
                .Paragraph(p => p.Bold("Reference / asset")).TextField("job-reference", width: global::OfficeIMO.OfficeLength.Points(300), value: "WO-018 / ROOM-04")
                .Paragraph(p => p.Bold("Assigned technician")).TextField("technician", width: global::OfficeIMO.OfficeLength.Points(300))
                .Paragraph(p => p.Bold("Planned work")).TextField("planned-work", width: global::OfficeIMO.OfficeLength.Points(450), height: global::OfficeIMO.OfficeLength.Points(44), value: "Inspect the meeting-room display and record the connection test.")
                .H2("Record the result", PdfAlign.Left, navy)
                .Paragraph(p => p.Bold("Work completed")).TextField("completed-work", width: global::OfficeIMO.OfficeLength.Points(450), height: global::OfficeIMO.OfficeLength.Points(44))
                .Paragraph(p => p.Bold("Outcome")).ChoiceField("outcome", new[] { "Awaiting work", "Completed", "Follow-up needed" }, value: "Awaiting work", width: 260)
                .Paragraph(p => p.Text("Handover acknowledged")).CheckBox("handover-acknowledged")
                .Paragraph(p => p.Bold("Follow-up owner and date")).TextField("follow-up", width: global::OfficeIMO.OfficeLength.Points(450))),
                new PdfCreateOptions { DefaultFontSize = 11 }.SetDefaultHeadingStyle(2, new PdfHeadingStyle { SpacingBefore = 14, SpacingAfter = 8 }))
                .Meta(title: "Maintenance job card", author: "OfficeIMO").Save(Path.Combine(folder, "example.pdf"));
        }
    }
}
