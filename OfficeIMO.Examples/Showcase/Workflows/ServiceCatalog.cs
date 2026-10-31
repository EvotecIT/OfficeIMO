using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Uses relative columns to present three service offerings in a single PDF.</summary>
internal static class ServiceCatalog {
    internal static void Create(string folder) {
        OfficeThemePalette palette = OfficeTheme.Default.Palette;
        OfficeColor muted = palette.TextMuted;
        PdfPanelStyle Note(OfficeColor accent) => new PdfPanelStyle {
            Background = palette.Background, BorderColor = accent,
            BorderWidth = OfficeLength.Points(0.75), PaddingX = OfficeLength.Points(10), PaddingY = OfficeLength.Points(8)
        };
        PdfDocument.Create(pdf => pdf.Content(content => {
            content.Title("Choose the support you need");
            content.Paragraph(p => p.FontSize(OfficeLength.Points(9)).Text("Northwind \u00B7 Service catalog"), defaultColor: muted);
            content.Subtitle("Three example offerings, with scope and handover expectations visible before work starts.");
            // The offerings and their notes sit in two rows, so the three notes start on one line
            // even when the lists above them differ in length.
            content.Row(row => {
                row.Gap(global::OfficeIMO.OfficeLength.Points(18));
                row.RelativeColumn(column => column
                    .H2("Essentials")
                    .Paragraph(p => p.Bold("Keep the service running."))
                    .Bullets(new[] { "Named support contact", "Monthly health summary", "Documented escalation path" }));
                row.RelativeColumn(column => column
                    .H2("Improvement")
                    .Paragraph(p => p.Bold("Reduce repeated problems."))
                    .Bullets(new[] { "Everything in Essentials", "Quarterly backlog review", "One improvement workshop" }));
                row.RelativeColumn(column => column
                    .H2("Transition")
                    .Paragraph(p => p.Bold("Prepare a confident handover."))
                    .Bullets(new[] { "Readiness assessment", "Recovery walkthrough", "Receiving-team training" }));
            });
            content.Row(row => {
                row.Gap(OfficeLength.Points(18));
                row.RelativeColumn(column => column
                    .PanelParagraph(p => p.Text("Best for a stable service with a small support queue."), Note(palette.Accent1)));
                row.RelativeColumn(column => column
                    .PanelParagraph(p => p.Text("Best for a service with recurring issues and a clear owner."), Note(palette.Accent3)));
                row.RelativeColumn(column => column
                    .PanelParagraph(p => p.Text("Best for a new service moving into regular operations."), Note(palette.Accent2)));
            });
            content.H2("Before choosing");
            content.Paragraph(p => p.Text("Confirm the service owner, support hours, and expected response process. These are sample offerings rather than a commercial contract."));
        }), new PdfCreateOptions { DefaultFontSize = OfficeLength.Points(10), ShowPageNumbers = true })
            .Meta(title: "Northwind service catalog", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
