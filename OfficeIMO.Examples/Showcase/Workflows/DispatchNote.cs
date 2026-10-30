using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a dispatch note with shipment details, item rows, and delivery instructions.</summary>
internal static class DispatchNote {
    internal static void Create(string folder) {
        global::OfficeIMO.OfficeColor navy = global::OfficeIMO.OfficeColor.FromRgb(23, 54, 93);
        var tableStyle = new PdfTableStyle {
            HeaderFill = navy, HeaderTextColor = global::OfficeIMO.OfficeColor.White, HeaderRowCount = 1,
            RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(241, 245, 249),
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225), BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(8), CellPaddingY = global::OfficeIMO.OfficeLength.Points(9), SpacingBefore = global::OfficeIMO.OfficeLength.Points(10), SpacingAfter = global::OfficeIMO.OfficeLength.Points(18)
        };
        string[][] items = {
            new[] { "Item", "Reference", "Quantity", "Package" },
            new[] { "Laptop workstation", "NW-LT-14", "8", "Cartons 1-4" },
            new[] { "USB-C docking station", "NW-DK-02", "8", "Carton 5" },
            new[] { "27-inch monitor", "NW-MN-27", "8", "Cartons 6-13" }
        };
        PdfDocument.Create(pdf => pdf.Content(content => content
            .H1("Dispatch note", PdfAlign.Left, navy)
            .Paragraph(p => p.Bold("DN-2026-0418").Text(" / Northwind equipment services"))
            .HR(global::OfficeIMO.OfficeLength.Points(1), navy, global::OfficeIMO.OfficeLength.Points(12), global::OfficeIMO.OfficeLength.Points(14))
            .H2("Deliver to", PdfAlign.Left, navy)
            .Paragraph(p => p.Text("Northwind Delivery Centre\nReceiving desk / Building B\nScheduled delivery: 18 September 2026"))
            .Table(items, style: tableStyle)
            .PanelParagraph(p => p.Bold("Receiving instructions\n").Text(
                "Check all 13 cartons against this note. Record visible damage before accepting the shipment."),
                new PdfPanelStyle {
                    Background = global::OfficeIMO.OfficeColor.FromRgb(232, 241, 251),
                    BorderColor = global::OfficeIMO.OfficeColor.FromRgb(165, 190, 225), PaddingX = 12, PaddingY = 12
                })
            .H2("Receipt", PdfAlign.Left, navy)
            .Paragraph(p => p.Text("Received by: ____________________     Date: ____________________"))
            .Paragraph(p => p.Text("Sample logistics document. Quantities and references are demonstration data."))),
            new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(11) })
            .Meta(title: "Dispatch note DN-2026-0418", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
