using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a two-page workshop pack with a shared header and numbered footer.</summary>
internal static class WorkshopPack {
    internal static void Create(string folder) {
        PdfColor accent = PdfColor.FromRgb(103, 58, 140);
        var options = new PdfOptions {
            DefaultFontSize = 11, ShowHeader = true, HeaderFormat = "NORTHWIND / SERVICE DESIGN",
            HeaderFontSize = 8, HeaderAlign = PdfAlign.Left,
            ShowPageNumbers = true, FooterFormat = "Workshop pack | {page} / {pages}",
            FooterFontSize = 8, FooterAlign = PdfAlign.Right
        };
        PdfDocument.Create(pdf => pdf.Content(content => content
            .H1("Service design workshop", PdfAlign.Left, accent)
            .Paragraph(p => p.Text("A half-day session to turn service feedback into a delivery plan."))
            .H2("What we will leave with", PdfAlign.Left, accent)
            .Bullets(new[] { "One agreed service journey", "Three prioritized improvements", "An owner and next step for each improvement" })
            .H2("Agenda", PdfAlign.Left, accent)
            .Table(new[] {
                new[] { "Time", "Session", "Output" },
                new[] { "09:00", "Current journey", "Shared map" },
                new[] { "09:45", "Friction and evidence", "Ranked pain points" },
                new[] { "10:45", "Options and trade-offs", "Three candidates" },
                new[] { "11:30", "Commitments", "Action list" }
            }, style: new PdfTableStyle {
                HeaderFill = accent, HeaderTextColor = PdfColor.White, HeaderRowCount = 1,
                CellPaddingX = 8, CellPaddingY = 8, RowStripeFill = PdfColor.FromRgb(247, 242, 251)
            })
            .PageBreak()
            .H1("Exercise sheet", PdfAlign.Left, accent)
            .H2("1. Describe one difficult moment", PdfAlign.Left, accent)
            .Paragraph(p => p.Text("Who is trying to do what, and where do they get stuck? Use an observed example."))
            .HR(0.5, PdfColor.FromRgb(203, 213, 225), 30, 30)
            .H2("2. Propose the smallest useful change", PdfAlign.Left, accent)
            .Paragraph(p => p.Text("Describe what the person will be able to do after the change."))
            .HR(0.5, PdfColor.FromRgb(203, 213, 225), 30, 30)
            .H2("3. Choose the first check", PdfAlign.Left, accent)
            .Paragraph(p => p.Text("What evidence would show that the change helped? Who will collect it?"))
            .HR(0.5, PdfColor.FromRgb(203, 213, 225), 30, 30)), options)
            .Meta(title: "Service design workshop pack", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
