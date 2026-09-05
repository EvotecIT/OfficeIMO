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
                  .typography { margin-top: 18px; font-family: Aptos, Arial, sans-serif; font-size: 20px; color: #336699; }
                  .decorated { font-weight: 700; font-style: italic; text-decoration-line: underline line-through; text-decoration-style: wavy; }
                  .small-caps { font-variant: small-caps; }
                </style>
                <main>
                  <h1>Quarterly status</h1>
                  <div class="summary">
                    <section class="card"><strong>API</strong><br>Consistent</section>
                    <section class="card"><strong>Renderer</strong><br>Direct</section>
                  </div>
                  <p class="typography"><span class="decorated">Styled output</span> H<sub>2</sub>O x<sup>2</sup> <span class="small-caps">Small caps</span></p>
                </main>
                """;

            var options = new HtmlToPdfOptions {
                PageSize = OfficePageSizes.A4,
                Margins = HtmlRenderMargins.All(32D),
                BackgroundColor = OfficeColor.White,
                Scale = 1.5D
            };

            string pdfPath = Path.Combine(folderPath, "Html11_DirectOutputs.pdf");
            string pngPath = Path.Combine(folderPath, "Html11_DirectOutputs.png");
            string svgPath = Path.Combine(folderPath, "Html11_DirectOutputs.svg");

            HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
            document.SaveAsPdf(pdfPath, options);
            document.SaveAsPng(pngPath, options);
            document.SaveAsSvg(svgPath, options);

            Console.WriteLine($"✓ Created: {pdfPath}");
            Console.WriteLine($"✓ Created: {pngPath}");
            Console.WriteLine($"✓ Created: {svgPath}");

            if (openPdf) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdfPath) { UseShellExecute = true });
            }
        }
    }
}
