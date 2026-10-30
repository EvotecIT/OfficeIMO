using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a dispatch note with shipment details, item rows, and delivery instructions.</summary>
internal static class DispatchNote {
    internal static void Create(string folder) {
        PdfColor muted = PdfColor.FromOfficeColor(OfficeTheme.Default.Palette.TextMuted);
        string[][] items = {
            new[] { "Item", "Reference", "Quantity", "Package" },
            new[] { "Laptop workstation", "NW-LT-14", "8", "Cartons 1–4" },
            new[] { "USB-C docking station", "NW-DK-02", "8", "Carton 5" },
            new[] { "27-inch monitor", "NW-MN-27", "8", "Cartons 6–13" }
        };
        PdfDocument.Create(pdf => pdf.Content(content => content
            .Title("Dispatch note")
            .Paragraph(p => p.Bold("DN-2026-0418").Text(" · Northwind equipment services"))
            .H2("Deliver to")
            .Paragraph(p => p.Text("Northwind Delivery Centre\nReceiving desk, Building B\nScheduled delivery: 18 September 2026"))
            .Table(items)
            .PanelParagraph(p => p.Bold("Receiving instructions\n").Text(
                "Check all 13 cartons against this note. Record visible damage before accepting the shipment."))
            .H2("Receipt")
            .Row(row => row.Gap(24)
                .RelativeColumn(column => column.Spacer(28).HR(0.75, muted, 0, 3).Paragraph(p => p.Text("Received by"), defaultColor: muted), 2)
                .RelativeColumn(column => column.Spacer(28).HR(0.75, muted, 0, 3).Paragraph(p => p.Text("Date"), defaultColor: muted), 1))
            .Paragraph(p => p.FontSize(9).Text("Sample logistics document. Quantities and references are demonstration data."), defaultColor: muted)),
            new PdfCreateOptions { ShowPageNumbers = true })
            .Meta(title: "Dispatch note DN-2026-0418", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
