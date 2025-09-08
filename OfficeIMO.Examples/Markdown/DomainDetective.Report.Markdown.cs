using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    /// <summary>
    /// Markdown counterpart to the Excel DomainDetective report example.
    /// Produces a readable, GitHub‑friendly report with TOC, callouts, tables, and per‑domain sections.
    /// </summary>
    internal static class DomainDetectiveReportMarkdown {
        private record DomainRow(
            string Domain,
            string MX,
            string SPF,
            string DKIM,
            string DMARC,
            string MTA_STS,
            string TLS_RPT,
            string Classification,
            string Findings
        );

        private record ScorePair(string Name, double Value);

        private record MailDomain(
            string Domain,
            string Classification,
            string Confidence,
            string[] ReceivingSignals,
            string[] SendingSignals,
            int Score,
            List<ScorePair> ScoreBreakdown,
            string Status,
            int WarningCount,
            int ErrorCount,
            string Summary,
            string[] Recommendations,
            string[] Positives,
            string[] References
        );

        public static void Example(string folderPath, bool _open) {
            Console.WriteLine("[*] Markdown - Domain Detective style report");
            string mdFolder = Path.Combine(folderPath, "Markdown");
            Directory.CreateDirectory(mdFolder);
            string mdPath = Path.Combine(mdFolder, "DomainDetective.Report.md");
            string htmlPath = Path.ChangeExtension(mdPath, ".html");

            // Demo data mirrors the Excel example
            var domains = new List<MailDomain> {
                new (
                    Domain: "evotec.pl",
                    Classification: "SendingAndReceiving",
                    Confidence: "High",
                    ReceivingSignals: new[]{"MX","TLS-RPT"},
                    SendingSignals: new[]{"SPF","DKIM","BIMI"},
                    Score: 8,
                    ScoreBreakdown: new(){ new("HasMX",2), new("HasNullMX",0), new("HasAorAAAA",0.5), new("EffectiveSPFSends",2) },
                    Status: "Warning",
                    WarningCount: 13,
                    ErrorCount: 0,
                    Summary: "SendingAndReceiving (High); recv 2; send 3",
                    Recommendations: new[]{"Enable DMARC enforcement","Rotate DKIM keys"},
                    Positives: new[]{"SPF present","DKIM present"},
                    References: new[]{
                        "https://datatracker.ietf.org/doc/html/rfc7208",
                        "https://datatracker.ietf.org/doc/html/rfc6376"
                    }
                ),
                new (
                    Domain: "evotec.xyz",
                    Classification: "SendingAndReceiving",
                    Confidence: "High",
                    ReceivingSignals: new[]{"MX","TLS-RPT"},
                    SendingSignals: new[]{"SPF","DKIM"},
                    Score: 7,
                    ScoreBreakdown: new(){ new("HasMX",2), new("HasNullMX",0), new("HasAorAAAA",0.5), new("EffectiveSPFSends",2) },
                    Status: "Warning",
                    WarningCount: 13,
                    ErrorCount: 0,
                    Summary: "SendingAndReceiving (High); recv 2; send 2",
                    Recommendations: new[]{"Consider BIMI"},
                    Positives: new[]{"SPF present","DKIM present"},
                    References: new[]{
                        "https://datatracker.ietf.org/doc/html/rfc7208",
                        "https://datatracker.ietf.org/doc/html/rfc6376"
                    }
                )
            };

            int totalWarnings = domains.Sum(d => d.WarningCount);
            int totalErrors = domains.Sum(d => d.ErrorCount);

            // Status helpers for Markdown tables (emoji badges keep it portable across renderers)
            static string Ok() => "🟢 OK";
            static string Warn() => "🟠 Warning";

            static string Slug(string text) {
                if (string.IsNullOrWhiteSpace(text)) return string.Empty;
                var sb = new System.Text.StringBuilder(text.Length);
                foreach (char ch in text.Trim().ToLowerInvariant()) {
                    if (ch >= 'a' && ch <= 'z') sb.Append(ch);
                    else if (ch >= '0' && ch <= '9') sb.Append(ch);
                    else if (char.IsWhiteSpace(ch) || ch == '_' || ch == '-' ) sb.Append('-');
                    // drop punctuation like '.' '/' ':' etc.
                }
                // collapse dashes
                var s = sb.ToString();
                while (s.Contains("--")) s = s.Replace("--", "-");
                return s.Trim('-');
            }

            var summaryRows = new List<DomainRow> {
                new($"[evotec.pl](#${Slug("evotec.pl")})", Warn(), Ok(), Warn(), Warn(), Warn(), Ok(), Warn(), "13 / 0"),
                new($"[evotec.xyz](#${Slug("evotec.xyz")})", Warn(), Ok(), Warn(), Warn(), Ok(), Ok(), Warn(), "13 / 0")
            };

            // Build Markdown document
            var md = MarkdownDoc.Create()
                .FrontMatter(new { title = "Domain Detective — Mail Classification", date = DateTimeOffset.Now.ToString("u") })
                .H1("Executive Summary")
                .TocAtTop("Contents", min: 1, max: 3)
                .H2("Overview")
                .P(p => p
                    .Text("This report summarizes the ")
                    .Bold("email security posture")
                    .Text($" for {domains.Count} domain(s). The table highlights the presence and status of key controls (MX, SPF, DKIM, DMARC, MTA-STS, TLS-RPT, Classification) and the count of warnings/errors detected. Total across all domains: ")
                    .Underline($"{totalWarnings} warning(s), {totalErrors} error(s)")
                    .Text("."))
                .Callout(totalErrors > 0 ? "warning" : "info", "Totals",
                    $"Warnings: {totalWarnings}\nErrors: {totalErrors}")
                .H2("Legend")
                .Table(t => t
                    .Headers("Status","Meaning")
                    .Row("🟢 OK","All checks passed or acceptable")
                    .Row("🟠 Warning","Requires attention; not blocking")
                    .Row("🔴 Error","Blocking or invalid configuration")
                    .AlignLeft(0,1))
                .H2("Domains")
                .Table(t => t
                    .Headers("Domain","MX","SPF","DKIM","DMARC","MTA-STS","TLS-RPT","Classification","Findings (W/E)")
                    .Rows(summaryRows.Select(r => (IReadOnlyList<string>)new[]{ r.Domain, r.MX, r.SPF, r.DKIM, r.DMARC, r.MTA_STS, r.TLS_RPT, r.Classification, r.Findings }))
                    .AlignLeft(0).AlignCenter(1,2,3,4,5,6,7).AlignRight(8))

                .H1("Background")
                .H2("SPF Overview")
                .P("Sender Policy Framework (SPF) lets a domain publish which mail servers are allowed to send on its behalf. Receiving servers can use this policy to detect and block spoofed email.")
                .P("SPF helps reduce impersonation and phishing by ensuring messages originate from authorized infrastructure. Together with DKIM and DMARC it provides robust protection against spoofing.")
                .Table(t => t
                    .Headers("Type","Meaning")
                    .Row("a","Authorize host A/AAAA addresses.")
                    .Row("mx","Authorize hosts listed as MX.")
                    .Row("ip4","Authorize IPv4 address or CIDR block.")
                    .Row("ip6","Authorize IPv6 address or CIDR block.")
                    .Row("include","Import another domain's SPF policy.")
                    .Row("exists","Authorize based on existence of DNS record.")
                    .Row("ptr","Authorize hosts by PTR (discouraged).")
                    .Row("redirect","Redirect evaluation to another domain.")
                    .Row("all","Catch‑all for remaining senders.")
                    .Row("version","SPF version token (v=spf1).")
                    .AlignLeft(0).AlignLeft(1))
                .H2("DKIM Overview")
                .P("DomainKeys Identified Mail (DKIM) uses a cryptographic signature to prove a message was authorized by the domain and not altered in transit.")
                .P("DKIM is one of the two mechanisms (alongside SPF) that can satisfy DMARC alignment. Strong keys and valid configuration improve deliverability and security.")
                .H2("DMARC Overview")
                .P("Domain-based Message Authentication, Reporting, and Conformance (DMARC) lets a domain specify policy for handling spoofed mail and receive feedback reports.")
                .P("DMARC reduces impersonation by requiring alignment of SPF and/or DKIM with the visible From domain, and enables receivers to report abuse.")
                .H2("MX Overview")
                .P("Mail Exchanger (MX) records direct where inbound mail should be delivered and influence reliability and resilience.")
                .P("Use multiple MX preferences and ensure consistency across name servers to avoid delivery issues.")
                .H2("Transport Policies")
                .P("MTA-STS lets a domain require TLS for inbound mail and publish a policy over HTTPS.")
                .P("TLS-RPT enables receivers to send reports about TLS negotiation failures to help identify misconfigurations.")
            ;

            // Per-domain sections
            foreach (var d in domains) {
                md.H1(d.Domain)
                  .H2("Overview")
                  .Table(t => t
                        .Headers("Key","Value")
                        .Row("Domain", d.Domain)
                        .Row("Classification", d.Classification)
                        .Row("Confidence", d.Confidence)
                        .Row("Status", d.Status)
                        .Row("Warnings", d.WarningCount.ToString())
                        .Row("Errors", d.ErrorCount.ToString())
                        .AlignLeft(0).AlignLeft(1))
                  .H2("Signals")
                  .Table(t => t
                        .Headers("Receiving","Sending")
                        .Row(string.Join(", ", d.ReceivingSignals), string.Join(", ", d.SendingSignals)))
                  .H2("Score")
                  .P($"Overall score: {d.Score} / 10")
                  .H2("Score Breakdown")
                  .Table(t => t
                        .Headers("Name","Value")
                        .Rows(d.ScoreBreakdown.Select(x => (IReadOnlyList<string>)new[]{ x.Name, x.Value.ToString("0.##") }))
                        .AlignLeft(0).AlignRight(1))
                  .H2("Legend")
                  .Table(t => t
                        .Headers("Status","Meaning")
                        .Row("🟢 OK","All checks passed or acceptable")
                        .Row("🟠 Warning","Requires attention; not blocking")
                        .Row("🔴 Error","Blocking or invalid configuration")
                        .AlignLeft(0,1));

                if (d.Recommendations.Length > 0) {
                    md.H2("Recommendations").Ul(d.Recommendations);
                }
                if (d.Positives.Length > 0) {
                    md.H2("Positives").Ul(d.Positives);
                }
                if (d.References.Length > 0) {
                    md.H2("References");
                    md.Ul(ul => {
                        foreach (var u in d.References) {
                            ul.ItemLink(u, u);
                        }
                    });
                }
            }

            // Emit Markdown and a styled HTML rendition for sharing
            File.WriteAllText(mdPath, md.ToMarkdown(), Encoding.UTF8);

            var htmlOptions = new HtmlOptions {
                Kind = HtmlKind.Document,
                Title = "Domain Detective — Mail Classification",
                Style = HtmlStyle.GithubAuto,
                CssDelivery = CssDelivery.Inline,
                IncludeAnchorLinks = false,
                ShowAnchorIcons = true,
                AnchorIcon = "🔗",
                CopyHeadingLinkOnClick = true,
                BackToTopLinks = true,
                BackToTopMinLevel = 1,
                BackToTopText = "Back to top",
                ThemeToggle = true
            };
            md.SaveHtml(htmlPath, htmlOptions);

            Console.WriteLine($"✓ Markdown saved: {mdPath}");
            Console.WriteLine($"✓ HTML saved:     {htmlPath}");
        }
    }
}
