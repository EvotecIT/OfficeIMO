using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Combines a decision panel, concise narrative, and action table in a weekly PDF brief.</summary>
internal static class WeeklyBrief {
    internal static void Create(string folder) {
        OfficeTheme theme = OfficeTheme.Default;
        OfficeColor muted = theme.Palette.TextMuted;
        OfficeColor warning = theme.Palette.Warning;
        string fonts = Path.Combine(AppContext.BaseDirectory, "Assets", "Fonts");
        var options = new PdfCreateOptions { ShowPageNumbers = true }
            .UseFontFamily(new PdfEmbeddedFontFamily("Carlito",
                File.ReadAllBytes(Path.Combine(fonts, "Carlito-Regular.ttf")),
                File.ReadAllBytes(Path.Combine(fonts, "Carlito-Bold.ttf"))))
            .UsePdfUa(PdfComplianceProfile.PdfUa1);
        PdfDocument.Create(pdf => pdf.Content(content => content
            .Title("The week in delivery")
            .Paragraph(p => p.FontSize(OfficeLength.Points(theme.TypeScale.Caption.ToPoints())).Text("Northwind · Week 38 · 2026"), defaultColor: muted)
            .PanelParagraph(p => p.Bold("Decision for this week\n").Text(
                "Approve the second pilot group after the recovery exercise is complete."))
            .H2("Progress")
            .Bullets(new[] {
                "The first pilot group completed 48 requests.",
                "The team reduced the oldest queue from 16 items to 5.",
                "The receiving team completed its first supported handover."
            })
            .H2("Watch closely")
            .Paragraph(p => p.Bold("At risk: ", warning).Text("two access issues remain open. Keep a named escalation contact available during the next pilot session."))
            .H2("Next actions")
            .Table(new[] {
                new[] { "Action", "Owner", "Due" },
                new[] { "Complete recovery exercise", "Operations", "Tuesday" },
                new[] { "Close access issues", "Engineering", "Wednesday" },
                new[] { "Review pilot expansion", "Product", "Friday" }
            })
            .Paragraph(p => p.FontSize(OfficeLength.Points(theme.TypeScale.Caption.ToPoints())).Text("Prepared from sample data. Keep the brief focused on decisions and actions."), defaultColor: muted)),
            options)
            .Meta(title: "The week in delivery", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
