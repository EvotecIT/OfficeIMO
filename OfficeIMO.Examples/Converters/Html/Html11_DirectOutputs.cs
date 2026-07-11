using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_Html11_DirectOutputs(string folderPath, bool openPdf) {
            Console.WriteLine("[*] HTML direct outputs: PDF, PNG, and SVG from one options object");

            const string html = """
                <style>
                  body { font-family: Arial; color: #203040; }
                  .summary { display: flex; gap: 12px; }
                  .card { padding: 12px; border: 1px solid #9fb3c8; border-radius: 8px; }
                </style>
                <main>
                  <h1>Quarterly status</h1>
                  <div class="summary">
                    <section class="card"><strong>API</strong><br>Consistent</section>
                    <section class="card"><strong>Renderer</strong><br>Direct</section>
                  </div>
                </main>
                """;

            var options = new HtmlPdfSaveOptions {
                PageSize = OfficePageSizes.A4,
                Margins = HtmlRenderMargins.All(32D),
                BackgroundColor = OfficeColor.White,
                Scale = 1.5D
            };

            string pdfPath = Path.Combine(folderPath, "Html11_DirectOutputs.pdf");
            string pngPath = Path.Combine(folderPath, "Html11_DirectOutputs.png");
            string svgPath = Path.Combine(folderPath, "Html11_DirectOutputs.svg");

            html.SaveAsPdf(pdfPath, options);
            html.SaveAsPng(pngPath, options);
            html.SaveAsSvg(svgPath, options);

            Console.WriteLine($"✓ Created: {pdfPath}");
            Console.WriteLine($"✓ Created: {pngPath}");
            Console.WriteLine($"✓ Created: {svgPath}");

            if (openPdf) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdfPath) { UseShellExecute = true });
            }
        }
    }
}
