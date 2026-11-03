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
        var options = new PdfCreateOptions { PageSetup = { Footer = { Show = true } } }
            .UseFontFamily(new PdfEmbeddedFontFamily("Carlito",
                File.ReadAllBytes(Path.Combine(fonts, "Carlito-Regular.ttf")),
                File.ReadAllBytes(Path.Combine(fonts, "Carlito-Bold.ttf"))))
            .UsePdfUa(PdfComplianceProfile.PdfUa1);
        PdfDocument.Create(pdf => pdf.Content(content => content
            .Title("The week in delivery")
            .AddParagraph(p => p.AddText("Northwind · Week 38 · 2026").SetFontSize(OfficeLength.Points(theme.TypeScale.Caption.ToPoints())), color: muted)
            .AddPanelParagraph(p => {
                p.AddText("Decision for this week\n").SetBold();
                p.AddText(
                "Approve the second pilot group after the recovery exercise is complete.");
            })
            .H2("Progress")
            .AddList(new[] {
                "The first pilot group completed 48 requests.",
                "The team reduced the oldest queue from 16 items to 5.",
                "The receiving team completed its first supported handover."
            })
            .H2("Watch closely")
            .AddParagraph(p => { p.AddText("At risk: ").SetBold().SetColor(warning); p.AddText("two access issues remain open. Keep a named escalation contact available during the next pilot session."); })
            .H2("Next actions")
            .AddTable(new[] {
                new[] { "Action", "Owner", "Due" },
                new[] { "Complete recovery exercise", "Operations", "Tuesday" },
                new[] { "Close access issues", "Engineering", "Wednesday" },
                new[] { "Review pilot expansion", "Product", "Friday" }
            })
            .AddParagraph(p => p.AddText("Prepared from sample data. Keep the brief focused on decisions and actions.").SetFontSize(OfficeLength.Points(theme.TypeScale.Caption.ToPoints())), color: muted)),
            options)
            .Meta(title: "The week in delivery", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
