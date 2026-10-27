using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a two-page workshop pack with a shared header and numbered footer.</summary>
internal static class WorkshopPack {
    internal static void Create(string folder) {
        PdfColor rule = PdfColor.FromOfficeColor(OfficeTheme.Default.TableStyle.BorderColor);
        var options = new PdfCreateOptions {
            ShowHeader = true, HeaderFormat = "Northwind · Service design",
            HeaderFontSize = 8.5, HeaderAlign = PdfAlign.Left,
            HeaderTextColor = PdfColor.FromOfficeColor(OfficeTheme.Default.Palette.TextMuted),
            ShowPageNumbers = true, FooterFontSize = 8.5, FooterAlign = PdfAlign.Right
        };
        PdfDocument.Create(pdf => pdf.Content(content => content
            .Title("Service design workshop")
            .Subtitle("A half-day session to turn service feedback into a delivery plan.")
            .H2("What we will leave with")
            .Bullets(new[] { "One agreed service journey", "Three prioritized improvements", "An owner and next step for each improvement" })
            .H2("Agenda")
            .Table(new[] {
                new[] { "Time", "Session", "Output" },
                new[] { "09:00", "Current journey", "Shared map" },
                new[] { "09:45", "Friction and evidence", "Ranked pain points" },
                new[] { "10:45", "Options and trade-offs", "Three candidates" },
                new[] { "11:30", "Commitments", "Action list" }
            })
            .PageBreak()
            .H1("Exercise sheet")
            .H3("1. Describe one difficult moment")
            .Paragraph(p => p.Text("Who is trying to do what, and where do they get stuck? Use an observed example."))
            .Spacer(130)
            .HR(0.5, rule, 0, 18)
            .H3("2. Propose the smallest useful change")
            .Paragraph(p => p.Text("Describe what the person will be able to do after the change."))
            .Spacer(130)
            .HR(0.5, rule, 0, 18)
            .H3("3. Choose the first check")
            .Paragraph(p => p.Text("What evidence would show that the change helped? Who will collect it?"))), options)
            .Meta(title: "Service design workshop pack", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
