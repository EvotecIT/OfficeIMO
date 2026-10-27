using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates a reviewable delivery-risk register with likelihood, impact, controls and owners.</summary>
    internal static class RiskAssessment {
        internal static void Create(string folder) {
            OfficeThemePalette palette = OfficeTheme.Default.Palette;
            PdfColor muted = PdfColor.FromOfficeColor(palette.TextMuted);
            // Status words carry the semantic colour in bold; the word itself keeps the meaning readable without colour.
            PdfTableCell Rating(string rating) => PdfTableCell.RichTextCell(new[] {
                PdfTextRun.Bolded(rating, PdfColor.FromOfficeColor(rating switch {
                    "High" => palette.Danger,
                    "Medium" => palette.Warning,
                    _ => palette.Success
                }))
            });
            PdfTableCell Text(string value) => PdfTableCell.TextCell(value);
            PdfDocument.Create(pdf => pdf.Content(content => content
                .Title("Pilot delivery risk review")
                .Paragraph(p => p.FontSize(9).Text("Northwind · Request portal · Review 03"), defaultColor: muted)
                .Paragraph(p => p.Text("Use likelihood and impact to prioritise discussion, then record a specific control and an owner. The ratings are illustrative judgments."))
                .H2("Open risks")
                .Table(new[] {
                    new[] { Text("Risk"), Text("Likelihood"), Text("Impact"), Text("Control"), Text("Owner") },
                    new[] { Text("Queue owner unavailable"), Rating("Medium"), Rating("Medium"), Text("Name a deputy"), Text("Operations") },
                    new[] { Text("Pilot access incomplete"), Rating("High"), Rating("Medium"), Text("Verify accounts before day one"), Text("Platform") },
                    new[] { Text("Recovery untested"), Rating("Low"), Rating("High"), Text("Complete restore exercise"), Text("Engineering") },
                    new[] { Text("Success measure unclear"), Rating("Medium"), Rating("High"), Text("Agree baseline and metric"), Text("Product") }
                })
                .H2("Decision before expansion")
                .PanelParagraph(p => p.Bold("Hold the next group until access and recovery are verified.").Text("\nThe sponsor records the acceptance decision; a low score alone is not approval."))
                .H2("Review discipline")
                .Bullets(new[] { "Revisit ratings when evidence changes.", "Keep the control, owner and verification date together.", "Close a risk only after the control is observed." })
                .Paragraph(p => p.FontSize(9).Text("Fictional project example. This is not a workplace safety assessment or compliance certification."), defaultColor: muted)),
                new PdfCreateOptions { ShowPageNumbers = true })
                .Meta(title: "Pilot delivery risk review", author: "OfficeIMO").Save(Path.Combine(folder, "example.pdf"));
        }
    }
}
