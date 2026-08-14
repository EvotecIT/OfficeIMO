using System.Net;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlInvoiceShowcase(string folderPath, bool openPdf) {
            Console.WriteLine("[*] HTML purchase report: data-bound rows, repeated headers, PDF, PNG, and SVG");

            IReadOnlyList<PurchaseLine> purchases = Enumerable.Range(0, 48)
                .Select(index => new PurchaseLine(
                    "SKU-" + index.ToString("D3"),
                    index == 11
                        ? "Precision calibrated field-service subscription with multilingual reporting, audit-ready evidence, extended retention, and regional compliance mapping"
                        : "Managed document service line " + index.ToString("D3"),
                    index % 4 + 1,
                    19.95m))
                .ToList();
            string html = BuildInvoiceHtml(purchases);

            var options = new HtmlPdfSaveOptions {
                PageSize = OfficePageSizes.A4,
                Margins = HtmlRenderMargins.All(24D),
                BackgroundColor = OfficeColor.White,
                Scale = 1.5D,
                PdfOptions = new PdfCore.PdfOptions {
                    FileVersion = PdfCore.PdfFileVersion.Pdf17,
                    ObjectSerializationMode = PdfCore.PdfObjectSerializationMode.ForwardOnly,
                    TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers
                }
            };

            string pdfPath = Path.Combine(folderPath, "HtmlInvoiceShowcase.pdf");
            string pngPath = Path.Combine(folderPath, "HtmlInvoiceShowcase.png");
            string svgPath = Path.Combine(folderPath, "HtmlInvoiceShowcase.svg");

            HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
            PdfCore.PdfSaveResult saved = document.SaveAsPdf(pdfPath, options).RequireSuccess();
            document.ToImage(options)
                .AsPng()
                .OnFileConflict(OfficeImageExportFileConflictPolicy.Replace)
                .Save(pngPath);
            document.ToImage(options)
                .AsSvg()
                .OnFileConflict(OfficeImageExportFileConflictPolicy.Replace)
                .Save(svgPath);

            Console.WriteLine($"    PDF: {pdfPath} ({saved.Serialization?.PageCount} pages, forward-only objects: {saved.Serialization?.IsForwardOnlyObjectSerialization})");
            Console.WriteLine($"    PNG: {pngPath}");
            Console.WriteLine($"    SVG: {svgPath}");

            if (openPdf) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdfPath) { UseShellExecute = true });
            }
        }

        private static string BuildInvoiceHtml(IReadOnlyList<PurchaseLine> purchases) {
            var rows = new StringBuilder(purchases.Count * 180);
            foreach (PurchaseLine purchase in purchases) {
                decimal total = purchase.Quantity * purchase.Rate;
                rows.Append("<tr><td><strong>")
                    .Append(WebUtility.HtmlEncode(purchase.Sku))
                    .Append("</strong><br>")
                    .Append(WebUtility.HtmlEncode(purchase.Description))
                    .Append("</td><td>")
                    .Append(purchase.Quantity)
                    .Append("</td><td class='amount'>$")
                    .Append(purchase.Rate.ToString("N2", System.Globalization.CultureInfo.InvariantCulture))
                    .Append("</td><td class='amount'>$")
                    .Append(total.ToString("N2", System.Globalization.CultureInfo.InvariantCulture))
                    .Append("</td></tr>");
            }

            decimal subtotal = purchases.Sum(purchase => purchase.Quantity * purchase.Rate);
            decimal tax = decimal.Round(subtotal * 0.08m, 2);
            decimal totalDue = subtotal + tax;
            return """
                <style>
                  @page {
                    size: A4;
                    margin: 16mm;
                    @top-center { content: "Northstar Works · INV-1042"; color:#52637a; font:9px Arial; }
                    @bottom-right { content: "Page " counter(page) " of " counter(pages); color:#64748b; font:8px Arial; }
                  }
                  body{margin:0;font:11px/1.35 Arial,sans-serif;color:#24324a;background:#fff}
                  main{background:#fff}
                  header,.party-grid,.totals{display:flex;justify-content:space-between;gap:24px}
                  .brand{color:#155eef;letter-spacing:.08em}.status{background:#e8f7ee;color:#176b3a;padding:5px 10px;border-radius:12px}
                  table{width:100%;border-collapse:collapse;table-layout:fixed;margin-top:14px}thead{display:table-header-group}th{background:#eef3fb;text-align:left}
                  th,td{padding:6px;border:1px solid #d8dfeb;vertical-align:top;overflow-wrap:anywhere}th:first-child,td:first-child{width:58%}.amount{text-align:right;white-space:nowrap}
                  .totals{break-inside:avoid;justify-content:flex-end;margin-top:14px}.total-card{width:230px;border-top:2px solid #155eef;padding-top:8px}
                  .total-card div{display:flex;justify-content:space-between}.cta{display:inline-block;background:#155eef;color:#fff;padding:8px 12px}
                </style>
                <main>
                  <header><div><strong class='brand'>NORTHSTAR WORKS</strong><h1>Invoice INV-1042</h1></div><div><span class='status'>Paid</span><p>Issued 11 July 2026<br>Due 25 July 2026</p></div></header>
                  <div class='party-grid'><section><h2>Bill to</h2><p><strong>Ada Lovelace</strong><br>12 Analytical Way<br>London</p></section>
                  <section><h2>From</h2><p>OfficeIMO Services<br>VAT PL-104200<br>Warsaw</p></section></div>
                  <table><thead><tr><th>Item</th><th>Qty</th><th class='amount'>Rate</th><th class='amount'>Total</th></tr></thead>
                  <tbody>ROWS</tbody></table>
                  <div class='totals'><div class='total-card'><div>Subtotal <strong>$SUBTOTAL</strong></div><div>Tax <strong>$TAX</strong></div><div>Total USD <strong>$TOTAL</strong></div></div></div>
                  <p><a class='cta' href='https://example.test/invoices/1042'>View invoice</a></p>
                </main>
                """
                .Replace("ROWS", rows.ToString())
                .Replace("SUBTOTAL", subtotal.ToString("N2", System.Globalization.CultureInfo.InvariantCulture))
                .Replace("TAX", tax.ToString("N2", System.Globalization.CultureInfo.InvariantCulture))
                .Replace("TOTAL", totalDue.ToString("N2", System.Globalization.CultureInfo.InvariantCulture));
        }

        private sealed record PurchaseLine(string Sku, string Description, int Quantity, decimal Rate);
    }
}
