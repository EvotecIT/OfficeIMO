using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Uses relative columns to present three service offerings in a single PDF.</summary>
internal static class ServiceCatalog {
    internal static void Create(string folder) {
        PdfColor navy = PdfColor.FromRgb(23, 54, 93);
        PdfDocument.Create(pdf => pdf.Content(content => {
            content.H1("Choose the support you need", PdfAlign.Left, navy);
            content.Paragraph(p => p.Text("NORTHWIND / SERVICE CATALOG"));
            content.Paragraph(p => p.Text("Three example offerings, with scope and handover expectations visible before work starts."));
            content.HR(1, navy, 12, 20);
            content.Row(row => {
                row.Gap(18);
                row.RelativeColumn(column => column
                    .H2("Essentials", PdfAlign.Left, navy)
                    .Paragraph(p => p.Bold("Keep the service running."))
                    .Bullets(new[] { "Named support contact", "Monthly health summary", "Documented escalation path" })
                    .PanelParagraph(p => p.Text("Best for a stable service with a small support queue."),
                        new PdfPanelStyle { Background = PdfColor.FromRgb(232, 241, 251), PaddingX = 10, PaddingY = 10 }));
                row.RelativeColumn(column => column
                    .H2("Improvement", PdfAlign.Left, navy)
                    .Paragraph(p => p.Bold("Reduce repeated problems."))
                    .Bullets(new[] { "Everything in Essentials", "Quarterly backlog review", "One improvement workshop" })
                    .PanelParagraph(p => p.Text("Best for a service with recurring issues and a clear owner."),
                        new PdfPanelStyle { Background = PdfColor.FromRgb(232, 246, 237), PaddingX = 10, PaddingY = 10 }));
                row.RelativeColumn(column => column
                    .H2("Transition", PdfAlign.Left, navy)
                    .Paragraph(p => p.Bold("Prepare a confident handover."))
                    .Bullets(new[] { "Readiness assessment", "Recovery walkthrough", "Receiving-team training" })
                    .PanelParagraph(p => p.Text("Best for a new service moving into regular operations."),
                        new PdfPanelStyle { Background = PdfColor.FromRgb(250, 241, 225), PaddingX = 10, PaddingY = 10 }));
            });
            content.H2("Before choosing", PdfAlign.Left, navy);
            content.Paragraph(p => p.Text("Confirm the service owner, support hours, and expected response process. These are sample offerings rather than a commercial contract."));
        }), new PdfOptions { DefaultFontSize = 10 })
            .Meta(title: "Northwind service catalog", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
