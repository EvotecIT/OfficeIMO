using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Authors a compact handbook with a generated table of contents and named sections.</summary>
internal static class NavigableHandbook {
    internal static void Create(string folder) {
        PdfColor navy = PdfColor.FromRgb(23, 54, 93);
        PdfDocument.Create(pdf => pdf.Content(content => content
            .H1("Community workshop handbook", PdfAlign.Left, navy)
            .Paragraph(p => p.Text("A practical reference for hosts, facilitators and volunteers."))
            .TableOfContents()
            .Section("Before participants arrive", section => section
                .Bullets(new[] { "Check access, seating and the sign-in desk.", "Test the demonstration equipment.", "Name the person who handles questions and changes." }))
            .Section("Run the session", section => section
                .Paragraph(p => p.Text("Explain the outcome, demonstrate one complete example and leave time for participants to try it themselves."))
                .Table(new[] {
                    new[] { "Segment", "Minutes", "Purpose" },
                    new[] { "Welcome", "10", "Set expectations and introduce the team" },
                    new[] { "Demonstration", "20", "Show one complete workflow" },
                    new[] { "Practice", "35", "Let participants apply the technique" },
                    new[] { "Review", "10", "Capture questions and next steps" }
                }, style: new PdfTableStyle { HeaderFill = navy, HeaderTextColor = PdfColor.White, HeaderRowCount = 1, CellPaddingX = 8, CellPaddingY = 8 }))
            .PageBreak()
            .Section("Close and follow up", section => section
                .Paragraph(p => p.Text("Finish with an outcome participants can keep: an example file, a checklist or a clear next action."))
                .H2("Before leaving", PdfAlign.Left, navy)
                .Bullets(new[] { "Collect shared equipment and remove temporary accounts.", "Record unanswered questions with named follow-up owners.", "Share the promised materials through the agreed channel." })
                .H2("Review the workshop", PdfAlign.Left, navy)
                .Paragraph(p => p.Text("Compare the planned outcome with what participants completed. Separate problems with the exercise from problems with the room or equipment."))
                .Paragraph(p => p.Text("Use the PDF table of contents or bookmarks to return to a section.")))),
            new PdfOptions { DefaultFontSize = 11, ShowPageNumbers = true, CreateOutlineFromHeadings = true })
            .Meta(title: "Community workshop handbook", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
