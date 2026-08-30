namespace OfficeIMO.Html.Pdf.Workbench;

public static class HtmlPdfWorkbenchTemplates {
    public static IReadOnlyList<HtmlPdfWorkbenchTemplate> All { get; } = Array.AsReadOnly(new[] {
        new HtmlPdfWorkbenchTemplate(
            "invoice",
            "Invoice",
            "A clean multi-section business document with a table and totals.",
            """
            <!doctype html><html lang="en"><head><meta charset="utf-8"><title>Invoice 1042</title></head><body>
            <main><header><div><p class="eyebrow">INVOICE</p><h1>Northwind Studio</h1><p>Design systems that ship.</p></div><div class="meta"><b>#1042</b><span>Issued 23 Aug 2026</span><span>Due 6 Sep 2026</span></div></header>
            <section class="addresses"><div><small>BILL TO</small><b>Adventure Works</b><span>Warsaw, Poland</span></div><div><small>FROM</small><b>Northwind Studio</b><span>Gdansk, Poland</span></div></section>
            <table><thead><tr><th>Service</th><th>Qty</th><th>Rate</th><th>Amount</th></tr></thead><tbody><tr><td>Design system audit</td><td>1</td><td>€2,400</td><td>€2,400</td></tr><tr><td>Component implementation</td><td>18</td><td>€180</td><td>€3,240</td></tr><tr><td>Accessibility review</td><td>1</td><td>€760</td><td>€760</td></tr></tbody></table>
            <section class="total"><span>Total</span><strong>€6,400</strong></section><footer>Thank you. Payment reference: INV-1042.</footer></main>
            </body></html>
            """,
            """
            :root{font-family:Inter,Arial,sans-serif;color:#172033;background:#eef2f7}*{box-sizing:border-box}body{margin:0;padding:44px}main{max-width:900px;margin:auto;background:white;padding:54px;border-radius:18px;box-shadow:0 24px 70px #1720331c}header,.addresses,.total{display:flex;justify-content:space-between;gap:32px}.eyebrow,small{color:#5b6b82;letter-spacing:.14em;font-size:12px}h1{font-size:34px;margin:7px 0}.meta,.addresses div{display:flex;flex-direction:column;gap:6px}.meta{text-align:right}.addresses{margin:52px 0 34px;padding:24px;background:#f5f7fb;border-radius:12px}table{width:100%;border-collapse:collapse}th,td{text-align:left;padding:16px 10px;border-bottom:1px solid #dfe5ee}th{font-size:12px;color:#5b6b82;text-transform:uppercase}th:not(:first-child),td:not(:first-child){text-align:right}.total{align-items:center;margin-left:auto;width:310px;padding:26px 10px}.total strong{font-size:28px;color:#3454d1}footer{border-top:1px solid #dfe5ee;padding-top:22px;color:#5b6b82}@page{size:A4;margin:12mm}
            """),
        new HtmlPdfWorkbenchTemplate(
            "accessible-report",
            "Accessible report",
            "Semantic headings, table headers, link text, and meaningful image alternatives.",
            """
            <!doctype html><html lang="en"><head><meta charset="utf-8"><title>Quarterly service report</title></head><body><main>
            <p class="kicker">QUARTERLY REPORT · Q3 2026</p><h1>Service reliability</h1><p class="lead">Availability improved while median response time remained below the target.</p>
            <section aria-labelledby="summary"><h2 id="summary">Executive summary</h2><div class="metrics"><article><b>99.98%</b><span>Availability</span></article><article><b>118 ms</b><span>Median latency</span></article><article><b>0</b><span>Critical incidents</span></article></div></section>
            <section aria-labelledby="regions"><h2 id="regions">Regional results</h2><table><caption>Availability by region</caption><thead><tr><th scope="col">Region</th><th scope="col">Availability</th><th scope="col">Change</th></tr></thead><tbody><tr><th scope="row">Europe</th><td>99.99%</td><td>+0.03</td></tr><tr><th scope="row">Americas</th><td>99.98%</td><td>+0.01</td></tr><tr><th scope="row">Asia Pacific</th><td>99.96%</td><td>+0.02</td></tr></tbody></table></section>
            <p>Read the <a href="https://example.com/methodology">measurement methodology</a> for definitions and exclusions.</p></main></body></html>
            """,
            """
            :root{font-family:Arial,sans-serif;color:#19221e}body{margin:0;padding:52px;background:#f3f7f4}main{max-width:920px;margin:auto;background:#fff;padding:58px;border-top:8px solid #167d55}.kicker{font-weight:700;color:#167d55;letter-spacing:.12em}h1{font-size:44px;margin:.25em 0}.lead{font-size:21px;color:#4d6157;max-width:680px}.metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin:30px 0}.metrics article{padding:24px;background:#eaf5ef;border-radius:10px}.metrics b,.metrics span{display:block}.metrics b{font-size:28px;color:#0d6744}.metrics span{margin-top:8px}h2{margin-top:42px}table{width:100%;border-collapse:collapse}caption{text-align:left;font-weight:700;margin-bottom:10px}th,td{padding:14px;border-bottom:1px solid #cddbd3;text-align:left}a{color:#075fa8}@page{size:A4;margin:14mm}
            """),
        new HtmlPdfWorkbenchTemplate(
            "print-poster",
            "Print poster",
            "A bleed-style visual sample for backgrounds, gradients, and CSS page sizing.",
            """
            <!doctype html><html lang="en"><head><meta charset="utf-8"><title>Field Notes</title></head><body><main><p>EDITION 07</p><h1>Field<br>Notes</h1><div class="rule"></div><h2>Observe closely.<br>Build deliberately.</h2><footer><span>OfficeIMO Studio</span><span>2026</span></footer></main></body></html>
            """,
            """
            @page{size:A4;margin:0}:root{font-family:Arial,sans-serif}*{box-sizing:border-box}body{margin:0}main{min-height:297mm;padding:24mm 20mm;color:#f8f4e8;background:radial-gradient(circle at 80% 20%,#ff9d4d 0 12%,transparent 13%),linear-gradient(135deg,#522bff,#171239 62%,#0b8f83);display:flex;flex-direction:column}p{letter-spacing:.3em;font-weight:700}h1{font-size:92px;line-height:.82;margin:32mm 0 12mm}.rule{height:3px;background:#ffcf71;width:42%}h2{font-size:30px;line-height:1.25;margin-top:12mm;font-weight:500}footer{display:flex;justify-content:space-between;margin-top:auto;border-top:1px solid #ffffff66;padding-top:8mm}
            """)
    });

    public static HtmlPdfWorkbenchTemplate Default => All[0];

    public static HtmlPdfWorkbenchTemplate Find(string id) =>
        All.FirstOrDefault(template => string.Equals(template.Id, id, StringComparison.Ordinal)) ?? Default;
}
