using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_PagedPurchaseTable_WrapsLongNamesRepeatsHeadersAndKeepsTotalsTogether() {
        const int rowCount = 84;
        var rows = new StringBuilder(rowCount * 240);
        decimal subtotal = 0m;
        for (int index = 0; index < rowCount; index++) {
            string marker = "SKU-" + index.ToString("D3");
            string description = index == 17
                ? "LongNameStart precision calibrated field-service subscription with multilingual reporting, audit-ready evidence, extended retention, regional compliance mapping, and LongNameEnd"
                : "Managed document service line " + index.ToString("D3");
            int quantity = index % 4 + 1;
            decimal amount = quantity * 19.95m;
            subtotal += amount;
            rows.Append("<tr><td><strong>")
                .Append(marker)
                .Append("</strong><br>")
                .Append(description)
                .Append("</td><td>")
                .Append(quantity)
                .Append("</td><td class='number'>$19.95</td><td class='number'>$")
                .Append(amount.ToString("N2", System.Globalization.CultureInfo.InvariantCulture))
                .Append("</td></tr>");
        }
        decimal tax = decimal.Round(subtotal * 0.08m, 2);
        decimal grandTotal = subtotal + tax;

        string html = """
            <style>
              @page {
                size: 5in 4in;
                margin: .42in .35in;
                @top-center { content: "Purchase report"; font: 9px Arial; color: #334155; }
                @bottom-right { content: "Page " counter(page) " of " counter(pages); font: 8px Arial; color: #64748b; }
              }
              body { margin: 0; font: 9px/1.35 Arial, sans-serif; color: #172033; }
              h1 { margin: 0 0 8px; font-size: 18px; }
              table { width: 100%; border-collapse: collapse; table-layout: fixed; }
              thead { display: table-header-group; }
              th { background: #dbeafe; text-align: left; }
              th, td { border: 1px solid #94a3b8; padding: 4px; vertical-align: top; overflow-wrap: anywhere; }
              th:first-child, td:first-child { width: 62%; }
              .number { text-align: right; white-space: nowrap; }
              .totals { break-inside: avoid; margin: 10px 0 0 auto; width: 180px; border-top: 2px solid #2563eb; padding-top: 5px; }
              .totals div { display: flex; justify-content: space-between; }
            </style>
            <main>
              <h1>Purchase statement PS-2048</h1>
              <table>
                <thead><tr><th>Item description</th><th>Qty</th><th class="number">Rate</th><th class="number">Amount</th></tr></thead>
                <tbody>ROWS</tbody>
              </table>
              <section class="totals">
                <div><span>SubtotalMarker</span><strong>$SUBTOTAL</strong></div>
                <div><span>TaxMarker</span><strong>$TAX</strong></div>
                <div><span>GrandTotalMarker</span><strong>$GRAND_TOTAL</strong></div>
              </section>
            </main>
            """
            .Replace("ROWS", rows.ToString())
            .Replace("SUBTOTAL", subtotal.ToString("N2", System.Globalization.CultureInfo.InvariantCulture))
            .Replace("TAX", tax.ToString("N2", System.Globalization.CultureInfo.InvariantCulture))
            .Replace("GRAND_TOTAL", grandTotal.ToString("N2", System.Globalization.CultureInfo.InvariantCulture));
        var options = new HtmlToPdfOptions {
            PageSize = new OfficePageSize(5D, 4D),
            Margins = HtmlRenderMargins.All(16D),
            PdfOptions = new PdfCore.PdfOptions {
                FileVersion = PdfCore.PdfFileVersion.Pdf17,
                ObjectSerializationMode = PdfCore.PdfObjectSerializationMode.ForwardOnly,
                TaggedStructureMode = PdfCore.PdfTaggedStructureMode.CatalogMarkers
            }
        };

        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        string renderedText = rendered.Text;
        IReadOnlyList<HtmlRenderPage> tablePages = rendered.Pages
            .Where(page => page.Visuals.OfType<HtmlRenderText>().Any(text => text.Text.StartsWith("SKU-", StringComparison.Ordinal)))
            .ToList();

        Assert.True(rendered.Pages.Count >= 7);
        Assert.NotEmpty(tablePages);
        Assert.All(tablePages, page =>
            Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Item description"));
        Assert.All(rendered.Pages, page => {
            Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Purchase report");
            Assert.Contains(page.Visuals.OfType<HtmlRenderText>(), text => text.Text == "Page " + page.PageNumber + " of " + rendered.Pages.Count);
        });
        for (int index = 0; index < rowCount; index++) {
            string marker = "SKU-" + index.ToString("D3");
            Assert.Equal(1, renderedText.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
        }
        Assert.Contains("LongNameStart", renderedText, StringComparison.Ordinal);
        Assert.Contains("LongNameEnd", renderedText, StringComparison.Ordinal);
        Assert.Equal(1, renderedText.Split(new[] { "SubtotalMarker" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(1, renderedText.Split(new[] { "TaxMarker" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(1, renderedText.Split(new[] { "GrandTotalMarker" }, StringSplitOptions.None).Length - 1);
        HtmlRenderPage totalsPage = Assert.Single(rendered.Pages, page =>
            page.Visuals.OfType<HtmlRenderText>().Any(text => text.Text == "GrandTotalMarker"));
        Assert.Contains(totalsPage.Visuals.OfType<HtmlRenderText>(), text => text.Text == "SubtotalMarker");
        Assert.Contains(totalsPage.Visuals.OfType<HtmlRenderText>(), text => text.Text == "TaxMarker");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed);

        using var output = new MemoryStream();
        PdfCore.PdfSaveResult save = document.SaveAsPdf(output, options).RequireSuccess();
        byte[] pdf = output.ToArray();
        PdfCore.PdfSerializationReport serialization = Assert.IsType<PdfCore.PdfSerializationReport>(save.Serialization);
        string pdfText = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Equal(rendered.Pages.Count, PdfCore.PdfInspector.Inspect(pdf).PageCount);
        Assert.Contains("LongNameStart", pdfText, StringComparison.Ordinal);
        Assert.Contains("LongNameEnd", pdfText, StringComparison.Ordinal);
        Assert.Equal(1, pdfText.Split(new[] { "GrandTotalMarker" }, StringSplitOptions.None).Length - 1);
        Assert.True(serialization.IsForwardOnlyObjectSerialization);
        Assert.True(serialization.UsesBoundedCompletedPayloadStores);
        Assert.False(serialization.IsForwardOnlyLayout);
    }
}
