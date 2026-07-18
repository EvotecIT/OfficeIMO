using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    public static IEnumerable<object[]> HtmlRenderingCorpusScenarioIds => CorpusCases
        .Select(item => new object[] { item.Id });

    [Fact]
    public void HtmlRenderingCorpus_CoversEveryPublishedMarketScenario() {
        Assert.Equal(
            HtmlMarketScenarioCatalog.All.Select(item => item.Id),
            CorpusCases.Select(item => item.Id));
    }

    [Theory]
    [MemberData(nameof(HtmlRenderingCorpusScenarioIds))]
    public void HtmlRenderingCorpus_ProvesSharedSceneImageAndSearchablePdf(string scenarioId) {
        HtmlRenderingCorpusCase scenario = CorpusCases.Single(item => item.Id == scenarioId);
        HtmlRenderOptions options = scenario.CreateOptions();

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(scenario.Html, options);

        Assert.Equal(scenario.Mode, rendered.Mode);
        Assert.Equal(scenario.ExpectedPageCount, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => {
            Assert.Equal(scenario.ExpectedSurfaceWidth, page.Width, 3);
            Assert.True(page.Height > 0D);
            Assert.True(page.Visuals.Count >= scenario.MinimumVisualCount);
        });
        Assert.True(rendered.Headings.Count >= scenario.MinimumHeadingCount);
        string logicalText = NormalizeCorpusWhitespace(rendered.Text);
        foreach (string marker in scenario.TextMarkers) Assert.Contains(NormalizeCorpusWhitespace(marker), logicalText, StringComparison.Ordinal);
        foreach (string code in scenario.DiagnosticCodes) {
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == code);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);
        if (scenario.LinkUri != null) {
            Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals), visual => visual.LinkUri == scenario.LinkUri);
        }

        OfficeDrawing firstPage = rendered.Pages[0].CreateDrawing();
        byte[] png = OfficeDrawingRasterRenderer.ToPng(firstPage, 0.5D, OfficeColor.White);
        string svg = OfficeDrawingSvgExporter.ToSvg(firstPage, 0.5D);
        Assert.True(png.Length > 100);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Take(8).ToArray());
        Assert.Contains("<svg", svg, StringComparison.Ordinal);
        foreach (string word in NormalizeCorpusWhitespace(scenario.TextMarkers[0]).Split(' ')) {
            Assert.Contains(word, svg, StringComparison.Ordinal);
        }

        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(scenario.Html).ToPdf(pdfOptions);
        PdfCore.PdfDocumentInfo pdfInfo = PdfCore.PdfInspector.Inspect(pdf);
        string pdfText = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Equal(scenario.ExpectedPageCount, pdfInfo.PageCount);
        string normalizedPdfText = NormalizeCorpusWhitespace(pdfText);
        foreach (string marker in scenario.TextMarkers) Assert.Contains(NormalizeCorpusWhitespace(marker), normalizedPdfText, StringComparison.Ordinal);
        if (scenario.LinkUri != null) Assert.Contains(scenario.LinkUri, pdfInfo.LinkUris);
    }

    private static IReadOnlyList<HtmlRenderingCorpusCase> CorpusCases { get; } = new[] {
        new HtmlRenderingCorpusCase(
            "invoice",
            HtmlRenderMode.Continuous,
            """
            <main><h1>Invoice INV-1042</h1><p>Bill to Ada Lovelace</p>
            <table style='width:100%;border-collapse:collapse'><thead><tr><th>Item</th><th>Total</th></tr></thead>
            <tbody><tr><td>Office suite</td><td>$420.00</td></tr></tbody></table>
            <p><a href='https://example.test/invoices/1042'>View invoice</a></p></main>
            """,
            new[] { "Invoice INV-1042", "Ada Lovelace", "$420.00" },
            linkUri: "https://example.test/invoices/1042"),
        new HtmlRenderingCorpusCase(
            "account-statement",
            HtmlRenderMode.Paged,
            """
            <section><h1>Account Statement</h1><p>Opening balance 1000.00</p>
            <table><tr><th>Date</th><th>Amount</th></tr><tr><td>2026-07-01</td><td>-25.00</td></tr></table></section>
            <section style='page-break-before:always'><h2>Statement Summary</h2><p>Closing balance 975.00</p></section>
            """,
            new[] { "Account Statement", "Closing balance 975.00" },
            expectedPageCount: 2),
        new HtmlRenderingCorpusCase(
            "quarterly-report",
            HtmlRenderMode.Continuous,
            """
            <article><h1>Quarterly Report</h1><p>Revenue increased by 18 percent.</p>
            <div style='display:flex;gap:12px'><section><h2>Revenue</h2><p>42000</p></section><section><h2>Margin</h2><p>18%</p></section></div>
            <table><tr><th>Region</th><th>Status</th></tr><tr><td>Europe</td><td>On plan</td></tr></table></article>
            """,
            new[] { "Quarterly Report", "Revenue increased", "Europe" },
            minimumHeadingCount: 3),
        new HtmlRenderingCorpusCase(
            "business-letter",
            HtmlRenderMode.Paged,
            """
            <article><h1>Project Confirmation</h1><p>11 July 2026</p><p>Dear Grace,</p>
            <p>We confirm the Atlas delivery schedule.</p><p>Sincerely,<br>OfficeIMO Team</p>
            <p><a href='mailto:office@example.test'>office@example.test</a></p></article>
            """,
            new[] { "Project Confirmation", "Dear Grace", "Atlas delivery schedule" },
            linkUri: "mailto:office@example.test"),
        new HtmlRenderingCorpusCase(
            "certificate",
            HtmlRenderMode.Paged,
            """
            <main style='border:8px double #234;padding:24px;text-align:center;background:linear-gradient(135deg,#ffffff,#dfefff)'>
            <h1>Certificate of Completion</h1><p>This certifies that</p><h2>Ada Lovelace</h2><p>completed the OfficeIMO program.</p></main>
            """,
            new[] { "Certificate of Completion", "Ada Lovelace", "OfficeIMO program" },
            minimumHeadingCount: 2),
        new HtmlRenderingCorpusCase(
            "product-catalog",
            HtmlRenderMode.Continuous,
            """
            <main><h1>Product Catalog</h1><div style='display:grid;grid-template-columns:repeat(2,1fr);gap:12px'>
            <article style='border:1px solid #999;padding:10px'><h2>Atlas</h2><p>Document automation</p><a href='https://example.test/products/atlas'>Details</a></article>
            <article style='border:1px solid #999;padding:10px'><h2>Nova</h2><p>Visual reporting</p></article></div></main>
            """,
            new[] { "Product Catalog", "Atlas", "Visual reporting" },
            linkUri: "https://example.test/products/atlas",
            minimumHeadingCount: 3),
        new HtmlRenderingCorpusCase(
            "legal-contract",
            HtmlRenderMode.Paged,
            """
            <article><h1>Services Agreement</h1><ol><li><strong>Scope.</strong> Deliver document tooling.</li>
            <li><strong>Term.</strong> Twelve months.</li></ol><p><a href='https://example.test/terms'>Referenced terms</a></p></article>
            """,
            new[] { "Services Agreement", "Deliver document tooling", "Twelve months" },
            linkUri: "https://example.test/terms"),
        new HtmlRenderingCorpusCase(
            "email-render",
            HtmlRenderMode.Continuous,
            """
            <main><h1>Action Required</h1><p style='color:#234'>Review the attached status update.</p>
            <p><a href='https://example.test/action'>Open action</a> <a href='javascript:alert(1)'>Unsafe action</a></p>
            <img src='file:///private/status.png' alt='Blocked status image'></main>
            """,
            new[] { "Action Required", "Review the attached status update", "Blocked status image" },
            linkUri: "https://example.test/action",
            diagnosticCodes: new[] { "ImageResourceRejectedByPolicy", "HyperlinkRejectedByPolicy" }),
        new HtmlRenderingCorpusCase(
            "dashboard-print",
            HtmlRenderMode.Continuous,
            """
            <main><h1>Operations Dashboard</h1><div style='display:grid;grid-template-columns:repeat(3,1fr);gap:10px'>
            <section style='background:#e8f2ff;padding:12px'><h2>Availability</h2><p>99.95%</p></section>
            <section style='background:#eef9ee;padding:12px'><h2>Jobs</h2><p>128</p></section>
            <section style='background:#fff4e5;padding:12px'><h2>Alerts</h2><p>2</p></section></div></main>
            """,
            new[] { "Operations Dashboard", "99.95%", "Alerts" },
            minimumHeadingCount: 4),
        new HtmlRenderingCorpusCase(
            "multilingual-bidi",
            HtmlRenderMode.Continuous,
            """
            <main><h1>Multilingual Summary</h1><p>English status: ready</p>
            <p dir='rtl'>שלום 123</p><p dir='rtl'>سلام 456</p><table><tr><th>Locale</th><th>Value</th></tr><tr><td>pl-PL</td><td>Zażółć</td></tr></table></main>
            """,
            new[] { "Multilingual Summary", "שלום 123", "سلام 456", "Zażółć" })
    };

    private static string NormalizeCorpusWhitespace(string value) {
        var result = new System.Text.StringBuilder(value.Length);
        bool pendingSpace = false;
        foreach (char character in value) {
            if (char.IsWhiteSpace(character)) {
                pendingSpace = result.Length > 0;
                continue;
            }
            if (pendingSpace) result.Append(' ');
            result.Append(character);
            pendingSpace = false;
        }
        return result.ToString();
    }

    private sealed class HtmlRenderingCorpusCase {
        internal HtmlRenderingCorpusCase(
            string id,
            HtmlRenderMode mode,
            string html,
            IReadOnlyList<string> textMarkers,
            int expectedPageCount = 1,
            string? linkUri = null,
            IReadOnlyList<string>? diagnosticCodes = null,
            int minimumVisualCount = 2,
            int minimumHeadingCount = 1) {
            Id = id;
            Mode = mode;
            Html = html;
            TextMarkers = textMarkers;
            ExpectedPageCount = expectedPageCount;
            LinkUri = linkUri;
            DiagnosticCodes = diagnosticCodes ?? Array.Empty<string>();
            MinimumVisualCount = minimumVisualCount;
            MinimumHeadingCount = minimumHeadingCount;
        }

        internal string Id { get; }
        internal HtmlRenderMode Mode { get; }
        internal string Html { get; }
        internal IReadOnlyList<string> TextMarkers { get; }
        internal int ExpectedPageCount { get; }
        internal string? LinkUri { get; }
        internal IReadOnlyList<string> DiagnosticCodes { get; }
        internal int MinimumVisualCount { get; }
        internal int MinimumHeadingCount { get; }
        internal double ExpectedSurfaceWidth => Mode == HtmlRenderMode.Paged ? 480D : 640D;

        internal HtmlRenderOptions CreateOptions() => new HtmlRenderOptions {
            Mode = Mode,
            ViewportWidth = 640D,
            PageSize = new OfficePageSize(5D, 4D),
            Margins = HtmlRenderMargins.All(24D),
            Scale = 0.5D,
            BackgroundColor = OfficeColor.White,
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        };
    }
}
