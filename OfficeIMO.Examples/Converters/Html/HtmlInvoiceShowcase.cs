using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlInvoiceShowcase(string folderPath, bool openPdf) {
            Console.WriteLine("[*] HTML invoice showcase: PDF, PNG, and SVG from one parsed document");

            const string html = """
                <style>
                  body{font:14px/1.4 Arial,sans-serif;color:#24324a;background:#f4f7fb}
                  main{max-width:720px;margin:16px auto;background:#fff;padding:28px;border:1px solid #d8dfeb}
                  header,.party-grid,.totals{display:flex;justify-content:space-between;gap:24px}
                  .brand{color:#155eef;letter-spacing:.08em}.status{background:#e8f7ee;color:#176b3a;padding:5px 10px;border-radius:12px}
                  table{width:100%;border-collapse:collapse;margin-top:20px}th{background:#eef3fb;text-align:left}
                  th,td{padding:9px;border-bottom:1px solid #d8dfeb}.amount{text-align:right}.totals{justify-content:flex-end;margin-top:14px}
                  .total-card{width:230px;border-top:2px solid #155eef;padding-top:8px}.cta{display:inline-block;background:#155eef;color:#fff;padding:8px 12px}
                </style>
                <main>
                  <header><div><strong class='brand'>NORTHSTAR WORKS</strong><h1>Invoice INV-1042</h1></div><div><span class='status'>Paid</span><p>Issued 11 July 2026<br>Due 25 July 2026</p></div></header>
                  <div class='party-grid'><section><h2>Bill to</h2><p><strong>Ada Lovelace</strong><br>12 Analytical Way<br>London</p></section>
                  <section><h2>From</h2><p>OfficeIMO Services<br>VAT PL-104200<br>Warsaw</p></section></div>
                  <table><thead><tr><th>Item</th><th>Qty</th><th class='amount'>Rate</th><th class='amount'>Total</th></tr></thead>
                  <tbody><tr><td>Office suite</td><td>1</td><td class='amount'>$420.00</td><td class='amount'>$420.00</td></tr>
                  <tr><td>PDF fidelity review</td><td>2</td><td class='amount'>$95.00</td><td class='amount'>$190.00</td></tr>
                  <tr><td>Document support</td><td>1</td><td class='amount'>$40.00</td><td class='amount'>$40.00</td></tr></tbody></table>
                  <div class='totals'><div class='total-card'><div>Subtotal <strong>$650.00</strong></div><div>Tax <strong>$52.00</strong></div><div>Total USD <strong>$702.00</strong></div></div></div>
                  <p><a class='cta' href='https://example.test/invoices/1042'>View invoice</a></p>
                </main>
                """;

            var options = new HtmlPdfSaveOptions {
                PageSize = OfficePageSizes.A4,
                Margins = HtmlRenderMargins.All(24D),
                BackgroundColor = OfficeColor.White,
                Scale = 1.5D
            };

            string pdfPath = Path.Combine(folderPath, "HtmlInvoiceShowcase.pdf");
            string pngPath = Path.Combine(folderPath, "HtmlInvoiceShowcase.png");
            string svgPath = Path.Combine(folderPath, "HtmlInvoiceShowcase.svg");

            HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
            document.SaveAsPdf(pdfPath, options);
            document.SaveAsPng(pngPath, options);
            document.SaveAsSvg(svgPath, options);

            Console.WriteLine($"    PDF: {pdfPath}");
            Console.WriteLine($"    PNG: {pngPath}");
            Console.WriteLine($"    SVG: {svgPath}");

            if (openPdf) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdfPath) { UseShellExecute = true });
            }
        }
    }
}
