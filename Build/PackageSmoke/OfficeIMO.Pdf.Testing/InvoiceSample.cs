using System;
using OfficeIMO.Pdf;

/// <summary>The invoice every framework consumer verifies. PDF_SNAPSHOT_SMOKE_TOTAL changes one table cell.</summary>
internal static class InvoiceSample {
    public static PdfDocument Create() {
        string total = Environment.GetEnvironmentVariable("PDF_SNAPSHOT_SMOKE_TOTAL") ?? "1,250.00";
        return PdfDocument.Create(pdf => pdf.Content(content => content
            .H1("Invoice 2026-118")
            .AddTable(new[] {
                new[] { "Item", "Amount" },
                new[] { "Support", "900.00" },
                new[] { "Total", total }
            })));
    }
}
