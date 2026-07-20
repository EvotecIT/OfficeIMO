using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Tests;

/// <summary>
/// Shared test-owned HTML rendering corpus used by HTML image and PDF artifact gates.
/// </summary>
internal static class HtmlRenderingCorpus {
    internal static IReadOnlyList<HtmlRenderingCorpusCase> All { get; } = new[] {
        new HtmlRenderingCorpusCase(
            "invoice",
            HtmlRenderMode.Continuous,
            """
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
            """,
            new[] { "Invoice INV-1042", "Ada Lovelace", "$702.00", "PDF fidelity review" },
            linkUri: "https://example.test/invoices/1042",
            minimumVisualCount: 20,
            minimumHeadingCount: 3,
            forbiddenDiagnosticCodes: new[] { "FlexLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "account-statement",
            HtmlRenderMode.Paged,
            """
            <style>
              @page{margin:42px 30px 38px;@top-left{content:"NORTHSTAR BANK"}@top-right{content:"Account • 1042"}@bottom-right{content:"Page " counter(page)}}
              body{font:11px/1.35 Arial,sans-serif;color:#1f2937}.summary{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
              .metric{background:#edf4ff;padding:9px;border-left:3px solid #2563eb}.metric strong{display:block;font-size:15px}
              table{width:100%;border-collapse:collapse;margin-top:12px}thead{display:table-header-group}th{background:#23395d;color:#fff}
              th,td{padding:6px;border-bottom:1px solid #d7deea;text-align:left}.money{text-align:right}.credit{color:#167044}
              .notice{border:1px solid #d7deea;background:#f8fafc;padding:10px}
            </style>
            <section>
              <h1>Account Statement</h1><p>Reporting period 1–31 July 2026 · IBAN PL00 1042 0000 0000</p>
              <div class='summary'><div class='metric'>Opening balance<strong>1000.00</strong></div><div class='metric'>Money out<strong>-245.00</strong></div><div class='metric'>Money in<strong>220.00</strong></div></div>
              <table><thead><tr><th>Date</th><th>Description</th><th>Reference</th><th class='money'>Amount</th></tr></thead><tbody>
                <tr><td>2026-07-01</td><td>Cloud subscription</td><td>CS-882</td><td class='money'>-25.00</td></tr>
                <tr><td>2026-07-04</td><td>Client transfer</td><td>ATLAS-7</td><td class='money credit'>+220.00</td></tr>
                <tr><td>2026-07-08</td><td>Insurance premium</td><td>INS-440</td><td class='money'>-80.00</td></tr>
                <tr><td>2026-07-12</td><td>Software license</td><td>LIC-901</td><td class='money'>-60.00</td></tr>
                <tr><td>2026-07-18</td><td>Office supplies</td><td>SUP-114</td><td class='money'>-40.00</td></tr>
                <tr><td>2026-07-24</td><td>Hosting</td><td>HOST-19</td><td class='money'>-40.00</td></tr>
              </tbody></table>
            </section>
            <section style='page-break-before:always'>
              <h2>Statement Summary</h2><div class='summary'><div class='metric'>Closing balance<strong>975.00</strong></div><div class='metric'>Transactions<strong>6</strong></div><div class='metric'>Currency<strong>EUR</strong></div></div>
              <h3>Account notes</h3><p class='notice'>Keep this statement with your accounting records. Amounts are shown in account currency.</p>
              <table><thead><tr><th>Category</th><th class='money'>Total</th></tr></thead><tbody><tr><td>Services</td><td class='money'>-125.00</td></tr><tr><td>Operations</td><td class='money'>-120.00</td></tr></tbody></table>
            </section>
            """,
            new[] { "Account Statement", "Closing balance 975.00", "Cloud subscription", "Account notes" },
            expectedPageCount: 2,
            minimumVisualCount: 24,
            minimumHeadingCount: 3,
            forbiddenDiagnosticCodes: new[] { "GridLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "quarterly-report",
            HtmlRenderMode.Continuous,
            """
            <style>
              body{font:13px/1.4 Arial,sans-serif;color:#25324a}article{max-width:760px;margin:auto}
              header{border-bottom:3px solid #5b4bdb}.filters{display:flex;gap:10px;align-items:end;background:#f5f6fb;padding:12px}
              label{display:flex;flex-direction:column;gap:3px}.kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:16px 0}
              .kpi{padding:12px;border-radius:6px;background:linear-gradient(135deg,#eef0ff,#fff);border:1px solid #d9dcf8}
              .kpi strong{display:block;font-size:20px;color:#4939c7}table{width:100%;border-collapse:collapse}th,td{padding:7px;border-bottom:1px solid #ddd;text-align:left}
              .good{color:#167044;font-weight:bold}
            </style>
            <article>
              <header><h1>Quarterly Report</h1><p>Revenue increased by 18 percent in Q2 2026.</p></header>
              <form class='filters'><label>Region<select id='report-region'><option selected>All regions</option><option>Europe</option></select></label>
                <label>Owner<input id='report-owner' value='Operations'></label><button id='apply-filter'>Apply filters</button></form>
              <div class='kpis'><section class='kpi'><h2>Revenue</h2><strong>42000</strong><small>+18% quarter over quarter</small></section>
                <section class='kpi'><h2>Margin</h2><strong>18%</strong><progress id='margin-progress' value='18' max='25'></progress></section>
                <section class='kpi'><h2>Pipeline</h2><strong>64 deals</strong><meter id='pipeline-meter' min='0' max='100' value='76'></meter></section></div>
              <h2>Regional performance</h2><table><thead><tr><th>Region</th><th>Revenue</th><th>Target</th><th>Status</th></tr></thead>
              <tbody><tr><td>Europe</td><td>21,300</td><td>20,000</td><td class='good'>On plan</td></tr><tr><td>Americas</td><td>14,200</td><td>15,000</td><td>Watch</td></tr><tr><td>Asia Pacific</td><td>6,500</td><td>6,000</td><td class='good'>On plan</td></tr></tbody></table>
            </article>
            """,
            new[] { "Quarterly Report", "Revenue increased", "All regions", "Apply filters", "72%", "76%", "Asia Pacific" },
            minimumVisualCount: 40,
            minimumHeadingCount: 5,
            forbiddenDiagnosticCodes: new[] { "FlexLayoutPending", "GridLayoutPending" },
            requiredVisualSources: new[] { "select#report-region", "input#report-owner", "button#apply-filter", "progress#margin-progress:value", "meter#pipeline-meter:value" }),
        new HtmlRenderingCorpusCase(
            "business-letter",
            HtmlRenderMode.Paged,
            """
            <style>
              @page{margin:48px;@bottom-left{content:"OfficeIMO Services · Warsaw"}}
              body{font:13px/1.55 Georgia,serif;color:#2f3540}article{max-width:560px;margin:auto}
              .letterhead{display:flex;justify-content:space-between;border-bottom:2px solid #1f5a8a;padding-bottom:14px}
              address{font-style:normal;color:#526070}.date{text-align:right}.signature{margin-top:28px}.contact{border-top:1px solid #ccd4dc;padding-top:10px}
            </style>
            <article><header class='letterhead'><div><strong>OFFICEIMO</strong><br><small>Document Engineering</small></div>
              <address>18 Builder Street<br>00-001 Warsaw<br>Poland</address></header>
              <p class='date'>11 July 2026</p><address>Grace Hopper<br>Compiler House<br>Arlington, Virginia</address>
              <h1>Project Confirmation</h1><p>Dear Grace,</p>
              <p>We confirm the Atlas delivery schedule and the agreed document-fidelity acceptance criteria.</p>
              <p>The first review package will include searchable Word, Excel, PowerPoint, and HTML-derived PDF artifacts.</p>
              <p class='signature'>Sincerely,<br><strong>OfficeIMO Team</strong><br>Document Platform Group</p>
              <p class='contact'>Questions: <a href='mailto:office@example.test'>office@example.test</a></p></article>
            """,
            new[] { "Project Confirmation", "Dear Grace", "Atlas delivery schedule", "PowerPoint", "Compiler House" },
            linkUri: "mailto:office@example.test",
            minimumVisualCount: 18,
            forbiddenDiagnosticCodes: new[] { "FlexLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "certificate",
            HtmlRenderMode.Paged,
            """
            <style>
              @page{margin:28px}body{font:14px/1.4 Georgia,serif;color:#21324d}
              main{border:8px double #234;padding:28px;text-align:center;background:linear-gradient(135deg,#ffffff 0%,#edf4ff 55%,#dceaff 100%);min-height:260px}
              .eyebrow{letter-spacing:.18em;color:#6f7d91}.name{font-size:30px;color:#1c4f80;margin:18px 0}
              .seal{display:inline-block;border:2px solid #b58a2e;border-radius:50%;padding:14px;background:radial-gradient(circle,#fff8da,#e8c65c)}
              .signatures{display:flex;justify-content:space-around;margin-top:24px}.line{border-top:1px solid #526070;padding-top:5px;min-width:130px}
            </style>
            <main><p class='eyebrow'>CERTIFIED ACHIEVEMENT</p><h1>Certificate of Completion</h1><p>This certifies that</p>
              <h2 class='name'>Ada Lovelace</h2><p>completed the <strong>OfficeIMO document fidelity program</strong><br>with distinction on 11 July 2026.</p>
              <div class='seal'>OFFICEIMO<br><strong>2026</strong></div>
              <div class='signatures'><div class='line'>Program Director</div><div class='line'>Document Engineer</div></div>
            </main>
            """,
            new[] { "Certificate of Completion", "Ada Lovelace", "OfficeIMO document fidelity program", "Program Director" },
            minimumVisualCount: 18,
            minimumHeadingCount: 2,
            forbiddenDiagnosticCodes: new[] { "FlexLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "product-catalog",
            HtmlRenderMode.Continuous,
            """
            <style>
              body{font:13px/1.4 Arial,sans-serif;color:#27344c;background:#f3f6fa}main{max-width:760px;margin:auto}
              .intro{display:flex;justify-content:space-between;align-items:end}.catalog{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}
              .card{position:relative;background:#fff;border:1px solid #d7dfeb;border-radius:8px;overflow:hidden;box-shadow:0 3px 9px rgba(20,40,70,.12)}
              .visual{height:62px;background:linear-gradient(120deg,#273b72,#4a7bd0)}.body{padding:13px}.badge{position:absolute;top:10px;right:10px;background:#fff;color:#273b72;padding:3px 7px;border-radius:10px}
              .price{font-size:18px;color:#155eef}.features{padding-left:18px}.details{display:inline-block;border:1px solid #155eef;padding:5px 9px;color:#155eef}
            </style>
            <main><header class='intro'><div><p>2026 COLLECTION</p><h1>Product Catalog</h1></div><p>Document platform components</p></header>
              <div class='catalog'>
                <article class='card'><div class='visual'></div><span class='badge'>Popular</span><div class='body'><h2>Atlas</h2><p>Document automation for teams.</p><ul class='features'><li>Office formats</li><li>Searchable PDF</li></ul><p class='price'>$49</p><a class='details' href='https://example.test/products/atlas'>Details</a></div></article>
                <article class='card'><div class='visual' style='background:linear-gradient(120deg,#166b5a,#49a87e)'></div><div class='body'><h2>Nova</h2><p>Visual reporting with reusable scenes.</p><ul class='features'><li>Charts</li><li>Dashboards</li></ul><p class='price'>$59</p></div></article>
                <article class='card'><div class='visual' style='background:linear-gradient(120deg,#7a3f16,#db8b42)'></div><div class='body'><h2>Orbit</h2><p>Batch conversion and compliance evidence.</p><ul class='features'><li>PDF/A checks</li><li>Diagnostics</li></ul><p class='price'>$79</p></div></article>
                <article class='card'><div class='visual' style='background:linear-gradient(120deg,#5d2b75,#a56ac0)'></div><div class='body'><h2>Pulse</h2><p>Operational document monitoring.</p><ul class='features'><li>Visual gates</li><li>Audit trails</li></ul><p class='price'>$39</p></div></article>
              </div>
            </main>
            """,
            new[] { "Product Catalog", "Atlas", "Visual reporting", "PDF/A checks", "Visual gates" },
            linkUri: "https://example.test/products/atlas",
            minimumVisualCount: 20,
            minimumHeadingCount: 5,
            forbiddenDiagnosticCodes: new[] { "GridLayoutPending", "PositionedLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "legal-contract",
            HtmlRenderMode.Paged,
            """
            <style>
              @page{margin:48px;@top-center{content:"SERVICES AGREEMENT · CONFIDENTIAL"}@bottom-center{content:"Initials: ______"}}
              body{font:11px/1.5 Georgia,serif;color:#20242a}article{max-width:600px;margin:auto}h1{text-align:center;letter-spacing:.05em}
              .parties{border:1px solid #b7bdc6;background:#f7f8fa;padding:10px}ol{counter-reset:clause;padding-left:22px}li{margin:8px 0}
              blockquote{border-left:3px solid #6b7280;margin:12px 0;padding-left:12px;color:#4b5563}.change{background:#fff5d6}
              .signatures{display:grid;grid-template-columns:1fr 1fr;gap:28px;margin-top:34px}.signature{border-top:1px solid #444;padding-top:5px}
            </style>
            <article><h1>Services Agreement</h1><p class='parties'>This agreement is between <strong>OfficeIMO Services</strong> (“Provider”) and <strong>Analytical Engines Ltd.</strong> (“Client”), effective 11 July 2026.</p>
              <ol><li><strong>Scope.</strong> Deliver document tooling for Word, Excel, PowerPoint, HTML, and searchable PDF output.
                <ol><li>Provide conversion diagnostics.</li><li>Maintain representative visual baselines.</li></ol></li>
                <li><strong>Term.</strong> Twelve months, unless renewed in writing.</li>
                <li><strong>Acceptance.</strong> The Client will review each delivery against the agreed fidelity corpus.</li>
                <li><strong>Change control.</strong> <span class='change'><del>Five</del> <ins>Ten</ins> business days’ notice is required.</span></li></ol>
              <blockquote>Confidential information remains the property of the disclosing party.</blockquote>
              <p>See the <a href='https://example.test/terms'>Referenced terms</a>, incorporated by reference.</p>
              <div class='signatures'><div class='signature'>Provider signature</div><div class='signature'>Client signature</div></div>
            </article>
            """,
            new[] { "Services Agreement", "Deliver document tooling", "Twelve months", "Ten business days", "Provider signature" },
            linkUri: "https://example.test/terms",
            minimumVisualCount: 26,
            forbiddenDiagnosticCodes: new[] { "GridLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "email-render",
            HtmlRenderMode.Continuous,
            """
            <table role='presentation' style='width:100%;background:#eef2f7;border-collapse:collapse'><tr><td style='padding:24px'>
              <table role='presentation' style='max-width:620px;margin:auto;background:#fff;border-collapse:collapse'>
                <tr><td style='background:#17345f;color:#fff;padding:18px'><strong style='font-size:18px'>NORTHSTAR OPERATIONS</strong></td></tr>
                <tr><td style='padding:22px;font:14px/1.5 Arial,sans-serif;color:#26364e'><p style='color:#59708f'>STATUS NOTIFICATION</p>
                  <h1>Action Required</h1><p>Review the attached status update before the 16:00 UTC deployment window.</p>
                  <div style='border-left:4px solid #f59e0b;background:#fff8e6;padding:12px'><strong>2 checks need attention</strong><br>Package signing and visual baseline approval.</div>
                  <p><a href='https://example.test/action' style='display:inline-block;background:#2563eb;color:#fff;padding:9px 14px'>Open action</a>
                  <a href='javascript:alert(1)'>Unsafe action</a></p>
                  <img src='file:///private/status.png' alt='Blocked status image' style='width:220px;height:80px'>
                  <p style='color:#667085;font-size:12px'>This message was generated automatically. Replying will not update the workflow.</p>
                </td></tr>
              </table>
            </td></tr></table>
            """,
            new[] { "Action Required", "Review the attached status update", "2 checks need attention", "Blocked status image" },
            linkUri: "https://example.test/action",
            diagnosticCodes: new[] { "ImageResourceRejectedByPolicy", "HyperlinkRejectedByPolicy" },
            minimumVisualCount: 20),
        new HtmlRenderingCorpusCase(
            "dashboard-print",
            HtmlRenderMode.Continuous,
            """
            <style>
              body{font:12px/1.35 Arial,sans-serif;color:#25324a;background:#f2f5f9}main{max-width:800px;margin:auto}
              header{display:flex;justify-content:space-between;align-items:end}.stamp{color:#667085}.cards{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}
              .card{background:#fff;border:1px solid #dae1eb;border-radius:7px;padding:12px}.metric{font-size:24px;font-weight:bold;color:#204f91}
              main>h2{font-size:16px;margin:18px 0 8px}.chart{display:flex;align-items:flex-end;gap:8px;height:110px;background:#fff;padding:14px;border:1px solid #dae1eb}
              .bar{flex:1;background:linear-gradient(180deg,#5b8def,#2656b0)}.legend{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:8px 0 16px}
              .incident{background:#fff4e5;border-left:4px solid #e58b23;padding:10px}
            </style>
            <main><header><div><p>PRINT SNAPSHOT</p><h1>Operations Dashboard</h1></div><p class='stamp'>20 July 2026 · 14:30 UTC</p></header>
              <div class='cards'><section class='card'><h2>Availability</h2><p class='metric'>99.95%</p><small>30-day service level</small></section>
                <section class='card'><h2>Jobs</h2><p class='metric'>128</p><small>12 currently running</small></section>
                <section class='card'><h2>Alerts</h2><p class='metric'>2</p><small>1 requires review</small></section></div>
              <h2>Documents processed</h2><div class='chart' aria-label='Seven day document volume'>
                <div class='bar' style='height:42%'></div><div class='bar' style='height:58%'></div><div class='bar' style='height:51%'></div>
                <div class='bar' style='height:76%'></div><div class='bar' style='height:68%'></div><div class='bar' style='height:91%'></div><div class='bar' style='height:84%'></div></div>
              <p class='legend'><span>Word 42%</span><span>Excel 31%</span><span>PowerPoint 17%</span><span>HTML 10%</span></p>
              <p class='incident'>Open incident: Two archived templates need font remapping.</p>
            </main>
            """,
            new[] { "Operations Dashboard", "99.95%", "Alerts", "Documents processed", "PowerPoint 17%", "Open incident" },
            minimumVisualCount: 38,
            minimumHeadingCount: 5,
            forbiddenDiagnosticCodes: new[] { "FlexLayoutPending", "GridLayoutPending" }),
        new HtmlRenderingCorpusCase(
            "multilingual-bidi",
            HtmlRenderMode.Continuous,
            """
            <style>
              body{font:14px/1.5 Arial,sans-serif;color:#26364e}main{max-width:720px;margin:auto}.summary{display:grid;grid-template-columns:1fr 1fr;gap:12px}
              .language{border:1px solid #d7dfeb;border-radius:6px;padding:12px}.rtl{direction:rtl;text-align:right;background:#f7f5ff}
              table{width:100%;border-collapse:collapse;margin-top:14px}th,td{padding:7px;border-bottom:1px solid #d7dfeb;text-align:left}
            </style>
            <main lang='en'><h1>Multilingual Summary</h1><p>English status: ready · Français: prêt · Deutsch: bereit</p>
              <div class='summary'><section class='language'><h2>Central Europe</h2><p lang='pl'>Zażółć gęślą jaźń</p><p lang='cs'>Příliš žluťoučký kůň</p></section>
                <section class='language rtl' dir='rtl'><h2>ملخص / סיכום</h2><p lang='he'>שלום 123 — מוכן</p><p lang='ar'>سلام 456 — جاهز</p></section></div>
              <table><thead><tr><th>Locale</th><th>Sample</th><th>Status</th></tr></thead><tbody>
                <tr><td>pl-PL</td><td>Zażółć</td><td>Gotowe</td></tr><tr><td>he-IL</td><td dir='rtl'>שלום 123</td><td>Ready</td></tr>
                <tr><td>ar-SA</td><td dir='rtl'>سلام 456</td><td>Ready</td></tr><tr><td>el-GR</td><td>Έγγραφο</td><td>Ready</td></tr>
              </tbody></table>
            </main>
            """,
            new[] { "Multilingual Summary", "שלום 123", "سلام 456", "Zażółć", "Příliš žluťoučký", "Έγγραφο" },
            minimumVisualCount: 28,
            minimumHeadingCount: 3,
            forbiddenDiagnosticCodes: new[] { "GridLayoutPending" })
    };
}

/// <summary>
/// Defines one stable HTML source and its observable cross-backend contract.
/// </summary>
internal sealed class HtmlRenderingCorpusCase {
    internal HtmlRenderingCorpusCase(
        string id,
        HtmlRenderMode mode,
        string html,
        IReadOnlyList<string> textMarkers,
        int expectedPageCount = 1,
        string? linkUri = null,
        IReadOnlyList<string>? diagnosticCodes = null,
        int minimumVisualCount = 2,
        int minimumHeadingCount = 1,
        IReadOnlyList<string>? forbiddenDiagnosticCodes = null,
        IReadOnlyList<string>? requiredVisualSources = null) {
        Id = id;
        Mode = mode;
        Html = html;
        TextMarkers = textMarkers;
        ExpectedPageCount = expectedPageCount;
        LinkUri = linkUri;
        DiagnosticCodes = diagnosticCodes ?? Array.Empty<string>();
        MinimumVisualCount = minimumVisualCount;
        MinimumHeadingCount = minimumHeadingCount;
        ForbiddenDiagnosticCodes = forbiddenDiagnosticCodes ?? Array.Empty<string>();
        RequiredVisualSources = requiredVisualSources ?? Array.Empty<string>();
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
    internal IReadOnlyList<string> ForbiddenDiagnosticCodes { get; }
    internal IReadOnlyList<string> RequiredVisualSources { get; }
    internal double ExpectedSurfaceWidth => Mode == HtmlRenderMode.Paged ? 8.27D * HtmlRenderOptions.CssPixelsPerInch : 640D;

    internal HtmlRenderOptions CreateOptions() => new HtmlRenderOptions {
        Mode = Mode,
        ViewportWidth = 640D,
        PageSize = new OfficePageSize(8.27D, 11.69D),
        Margins = HtmlRenderMargins.All(40D),
        Scale = 0.5D,
        BackgroundColor = OfficeColor.White,
        UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
    };
}
