using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Uses relative columns to present three service offerings in a single PDF.</summary>
internal static class ServiceCatalog {
    internal static void Create(string folder) {
        OfficeThemePalette palette = OfficeTheme.Default.Palette;
        PdfColor muted = PdfColor.FromOfficeColor(palette.TextMuted);
        PdfPanelStyle Note(OfficeColor accent) => new PdfPanelStyle {
            Background = PdfColor.FromOfficeColor(palette.Background), BorderColor = PdfColor.FromOfficeColor(accent),
            BorderWidth = 0.75, PaddingX = 10, PaddingY = 8
        };
        PdfDocument.Create(pdf => pdf.Content(content => {
            content.Title("Choose the support you need");
            content.Paragraph(p => p.FontSize(9).Text("Northwind \u00B7 Service catalog"), defaultColor: muted);
            content.Subtitle("Three example offerings, with scope and handover expectations visible before work starts.");
            content.Row(row => {
                row.Gap(18);
                row.RelativeColumn(column => column
                    .H2("Essentials")
                    .Paragraph(p => p.Bold("Keep the service running."))
                    .Bullets(new[] { "Named support contact", "Monthly health summary", "Documented escalation path" })
                    .PanelParagraph(p => p.Text("Best for a stable service with a small support queue."),
                        Note(palette.Accent1)));
                row.RelativeColumn(column => column
                    .H2("Improvement")
                    .Paragraph(p => p.Bold("Reduce repeated problems."))
                    .Bullets(new[] { "Everything in Essentials", "Quarterly backlog review", "One improvement workshop" })
                    .PanelParagraph(p => p.Text("Best for a service with recurring issues and a clear owner."),
                        Note(palette.Accent3)));
                row.RelativeColumn(column => column
                    .H2("Transition")
                    .Paragraph(p => p.Bold("Prepare a confident handover."))
                    .Bullets(new[] { "Readiness assessment", "Recovery walkthrough", "Receiving-team training" })
                    .PanelParagraph(p => p.Text("Best for a new service moving into regular operations."),
                        Note(palette.Accent2)));
            });
            content.H2("Before choosing");
            content.Paragraph(p => p.Text("Confirm the service owner, support hours, and expected response process. These are sample offerings rather than a commercial contract."));
        }), new PdfCreateOptions { DefaultFontSize = 10, ShowPageNumbers = true })
            .Meta(title: "Northwind service catalog", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
