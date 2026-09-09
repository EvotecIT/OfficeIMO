using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Combines a decision panel, concise narrative, and action table in a weekly PDF brief.</summary>
internal static class WeeklyBrief {
    internal static void Create(string folder) {
        PdfColor navy = PdfColor.FromRgb(23, 54, 93);
        PdfDocument.Create(pdf => pdf.Content(content => content
            .H1("The week in delivery", PdfAlign.Left, navy)
            .Paragraph(p => p.Text("NORTHWIND / WEEK 38 / 2026"))
            .PanelParagraph(p => p.Bold("Decision for this week\n").Text(
                "Approve the second pilot group after the recovery exercise is complete."),
                new PdfPanelStyle {
                    Background = PdfColor.FromRgb(232, 246, 237),
                    BorderColor = PdfColor.FromRgb(117, 170, 135), PaddingX = 12, PaddingY = 12
                })
            .H2("Progress", PdfAlign.Left, navy)
            .Bullets(new[] {
                "The first pilot group completed 48 requests.",
                "The team reduced the oldest queue from 16 items to 5.",
                "The receiving team completed its first supported handover."
            })
            .H2("Watch closely", PdfAlign.Left, navy)
            .Paragraph(p => p.Text("Two access issues remain open. Keep a named escalation contact available during the next pilot session."))
            .H2("Next actions", PdfAlign.Left, navy)
            .Table(new[] {
                new[] { "Action", "Owner", "Due" },
                new[] { "Complete recovery exercise", "Operations", "Tuesday" },
                new[] { "Close access issues", "Engineering", "Wednesday" },
                new[] { "Review pilot expansion", "Product", "Friday" }
            }, style: new PdfTableStyle {
                HeaderFill = navy, HeaderTextColor = PdfColor.White, HeaderRowCount = 1,
                RowStripeFill = PdfColor.FromRgb(241, 245, 249), CellPaddingX = 8, CellPaddingY = 8
            })
            .Paragraph(p => p.Text("Prepared from sample data. Keep the brief focused on decisions and actions."))),
            new PdfOptions { DefaultFontSize = 11 })
            .Meta(title: "The week in delivery", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
