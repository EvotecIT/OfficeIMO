using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates a reviewable delivery-risk register with likelihood, impact, controls and owners.</summary>
    internal static class RiskAssessment {
        internal static void Create(string folder) {
            var navy = global::OfficeIMO.OfficeColor.FromRgb(23, 54, 93);
            PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Pilot delivery risk review", PdfAlign.Left, navy)
                .Paragraph(p => p.Text("NORTHWIND / REQUEST PORTAL / REVIEW 03"))
                .Paragraph(p => p.Text("Use likelihood and impact to prioritise discussion, then record a specific control and an owner. The ratings are illustrative judgments."))
                .H2("Open risks", PdfAlign.Left, navy)
                .Table(new[] {
                    new[] { "Risk", "Likelihood", "Impact", "Control / owner" },
                    new[] { "Queue owner unavailable", "Medium", "Medium", "Name a deputy / Operations" },
                    new[] { "Pilot access incomplete", "High", "Medium", "Verify accounts before day one / Platform" },
                    new[] { "Recovery untested", "Low", "High", "Complete restore exercise / Engineering" },
                    new[] { "Success measure unclear", "Medium", "High", "Agree baseline and metric / Product" }
                }, style: new PdfTableStyle { HeaderFill = navy, HeaderTextColor = global::OfficeIMO.OfficeColor.White, HeaderRowCount = 1, CellPaddingX = global::OfficeIMO.OfficeLength.Points(8), CellPaddingY = global::OfficeIMO.OfficeLength.Points(10), RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(241, 245, 249) })
                .H2("Decision before expansion", PdfAlign.Left, navy)
                .PanelParagraph(p => p.Bold("Hold the next group until access and recovery are verified.").Text("\nThe sponsor records the acceptance decision; a low score alone is not approval."),
                    new PdfPanelStyle { Background = global::OfficeIMO.OfficeColor.FromRgb(255, 244, 214), PaddingX = 12, PaddingY = 12 })
                .H2("Review discipline", PdfAlign.Left, navy)
                .Bullets(new[] { "Revisit ratings when evidence changes.", "Keep the control, owner and verification date together.", "Close a risk only after the control is observed." })
                .Paragraph(p => p.Text("Fictional project example. This is not a workplace safety assessment or compliance certification."))),
                new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(11), ShowPageNumbers = true }.SetDefaultHeadingStyle(2, new PdfHeadingStyle { SpacingBefore = 14, SpacingAfter = 8 }))
                .Meta(title: "Pilot delivery risk review", author: "OfficeIMO").Save(Path.Combine(folder, "example.pdf"));
        }
    }
}
