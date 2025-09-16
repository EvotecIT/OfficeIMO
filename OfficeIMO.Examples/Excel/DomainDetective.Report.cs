using System;
using System.Collections.Generic;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using System.Linq;

namespace OfficeIMO.Examples.Excel {
    public static class DomainDetectiveReport {
        // Minimal demo models simulating your output
        private record ScorePair(string Name, double Value);
        private record MailDomainClassificationResult(
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

        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Domain Detective style report");
            string filePath = System.IO.Path.Combine(folderPath, "DomainDetective.Report.xlsx");

            var data = new List<MailDomainClassificationResult> {
                new (
                    Domain: "evotec.pl",
                    Classification: "SendingAndReceiving",
                    Confidence: "High",
                    ReceivingSignals: new[]{"MX","TLS-RPT"},
                    SendingSignals: new[]{"SPF","DKIM","BIMI"},
                    Score: 8,
                    ScoreBreakdown: new(){ new("HasMX",2), new("HasNullMX",0), new("HasAorAAAA",0.5), new("EffectiveSPFSends",2) },
                    Status: "Warning",
                    WarningCount: 6,
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
                    WarningCount: 4,
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

            using (var doc = ExcelDocument.Create(filePath)) {
                // Document properties via fluent Info
                doc.AsFluent().Info(i => i
                    .Title("Domain Detective — Mail Classification")
                    .Author("OfficeIMO")
                    .LastModifiedBy("OfficeIMO")
                    .Company("Evotec")
                    .Application("OfficeIMO.Excel")
                    .Keywords("email,security,classification,excel")
                ).End();
                // Sheet 1: Summary of all domains as a table (SheetComposer)
                var composer = new SheetComposer(doc, "Summary");
                composer.Title("Domain Detective — Mail Classification Summary");
                // Totals callout (parity with Markdown)
                int totalWarnings = data.Sum(d => d.WarningCount);
                int totalErrors = data.Sum(d => d.ErrorCount);
                composer.Callout(totalErrors > 0 ? "warning" : "info", "Totals",
                    $"Warnings: {totalWarnings}. Errors: {totalErrors}.");

                var summaryRange = composer.TableFrom(data, title: "Domains", configure: opts => {
                    // Map ScoreBreakdown list into dynamic columns using Name as header and Value as cell
                    opts.CollectionMapColumns[nameof(MailDomainClassificationResult.ScoreBreakdown)] = new CollectionColumnMapping { KeyProperty = nameof(ScorePair.Name), ValueProperty = nameof(ScorePair.Value) };
                    // Make headers nice
                    opts.HeaderCase = HeaderCase.Title;
                    opts.HeaderPrefixTrimPaths = new[] { nameof(MailDomainClassificationResult.ScoreBreakdown) + "." };
                    opts.NullPolicy = NullPolicy.EmptyString;
                    // Keep important fields at the front of the table; the rest follow automatically
                    opts.PinnedFirst = new[]
                    {
                        nameof(MailDomainClassificationResult.Domain),
                        nameof(MailDomainClassificationResult.Classification),
                        nameof(MailDomainClassificationResult.Confidence),
                        nameof(MailDomainClassificationResult.Status),
                        nameof(MailDomainClassificationResult.Score),
                        nameof(MailDomainClassificationResult.WarningCount),
                        nameof(MailDomainClassificationResult.ErrorCount),
                        nameof(MailDomainClassificationResult.ReceivingSignals),
                        nameof(MailDomainClassificationResult.SendingSignals),
                        nameof(MailDomainClassificationResult.Summary)
                    };
                    // Keep long text columns out of the summary table
                    opts.Ignore = new[] { nameof(MailDomainClassificationResult.Recommendations), nameof(MailDomainClassificationResult.Positives), nameof(MailDomainClassificationResult.References) };
                }, style: TableStyle.TableStyleMedium9, visuals: viz => {
                    viz.IconSetColumns.Add("Score");
                    (viz.TextBackgrounds["Status"], viz.BoldByText["Status"]) = StatusPalettes.Default;
                });
                // Emphasize statuses (bold for Error/Warning)
                // Also request bolding via visuals fallback (works even when header row isn't the first row)
                // Note: we keep a direct call as a best-effort attempt; visuals will ensure it applies.
                try { composer.Sheet.ColumnStyleByHeader("Status").BoldByTextSet(new HashSet<string>(StringComparer.OrdinalIgnoreCase){"Error","Warning"}); } catch { }

                // Make Summary sheet print nicely
                composer.Sheet.SetGridlinesVisible(false);
                composer.Sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0);
                doc.SetPrintArea(composer.Sheet, summaryRange);
                // Legend
                composer.Section("Legend");
                var legendHeaderRow = composer.CurrentRow;
                composer.Sheet.Cell(legendHeaderRow, 1, "Status");
                composer.Sheet.CellBold(legendHeaderRow, 1, true);
                composer.Sheet.CellBackground(legendHeaderRow, 1, "#F2F2F2");
                composer.Sheet.Cell(legendHeaderRow, 2, "Meaning");
                composer.Sheet.CellBold(legendHeaderRow, 2, true);
                composer.Sheet.CellBackground(legendHeaderRow, 2, "#F2F2F2");
                var palette = StatusPalettes.Default;
                // OK / Success
                composer.Sheet.Cell(legendHeaderRow + 1, 1, "OK");
                if (palette.FillHexMap.TryGetValue("OK", out var okHex)) composer.Sheet.CellBackground(legendHeaderRow + 1, 1, okHex);
                composer.Sheet.Cell(legendHeaderRow + 1, 2, "All checks passed or acceptable");
                // Warning
                composer.Sheet.Cell(legendHeaderRow + 2, 1, "Warning");
                if (palette.FillHexMap.TryGetValue("Warning", out var warnHex)) composer.Sheet.CellBackground(legendHeaderRow + 2, 1, warnHex);
                composer.Sheet.Cell(legendHeaderRow + 2, 2, "Requires attention; not blocking");
                // Error
                composer.Sheet.Cell(legendHeaderRow + 3, 1, "Error");
                if (palette.FillHexMap.TryGetValue("Error", out var errHex)) composer.Sheet.CellBackground(legendHeaderRow + 3, 1, errHex);
                composer.Sheet.Cell(legendHeaderRow + 3, 2, "Blocking or invalid configuration");
                composer.Spacer();
                // Header/footer via fluent builder
                var logoPath = System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", "OfficeIMO.png");
                byte[]? logo = System.IO.File.Exists(logoPath) ? System.IO.File.ReadAllBytes(logoPath) : null;
                composer.HeaderFooter(h =>
                {
                    h.Center("Domain Detective").Right("Page &P of &N");
                    if (logo != null) h.CenterImage(logo, widthPoints: 96, heightPoints: 32);
                });
                // Back links to TOC on all sheets
                doc.AddBackLinksToToc();
                // Horizontal band: avoid full-sheet auto-fit to prevent column width fights
                composer.Finish(autoFitColumns: false);

                // One sheet per domain with richer layout
                foreach (var d in data) {
                    var rs = new SheetComposer(doc, d.Domain);
                    rs.Title($"Mail Classification — {d.Domain}", d.Summary)
                      // Per-domain status callout
                      .Callout(d.ErrorCount > 0 ? "error" : (string.Equals(d.Status, "Warning", StringComparison.OrdinalIgnoreCase) ? "warning" : "info"),
                          "Status",
                          $"Status: {d.Status}; Findings: {d.WarningCount} warning(s), {d.ErrorCount} error(s).")
                      .SectionWithAnchor("Overview")
                      .DefinitionList(new (string, object?)[] {
                          ("Domain", d.Domain),
                          ("Classification", d.Classification),
                          ("Confidence", d.Confidence),
                          ("Status", d.Status),
                          ("Warnings", d.WarningCount),
                          ("Errors", d.ErrorCount)
                      }, columns: 3)
                      .SectionWithAnchor("Signals")
                      .PropertiesGrid(new (string, object?)[] {
                          ("Receiving", string.Join(", ", d.ReceivingSignals)),
                          ("Sending", string.Join(", ", d.SendingSignals))
                      }, columns: 2)
                      .Score("Score", d.Score)
                      .SectionWithAnchor("Score Breakdown");

                    rs.TableFrom(d.ScoreBreakdown, title: null, configure: o => {
                        o.Columns = new[] { nameof(ScorePair.Name), nameof(ScorePair.Value) };
                        o.HeaderCase = HeaderCase.Title;
                    }, visuals: v => {
                        // Example of explicit icon thresholds for demo purposes
                        v.IconSets["Value"] = new IconSetOptions {
                            IconSet = DocumentFormat.OpenXml.Spreadsheet.IconSetValues.ThreeSymbols,
                            ShowValue = true,
                            ReverseOrder = false,
                            PercentThresholds = new double[] { 0, 60, 85 } // 0-60 red, 60-85 yellow, 85-100 green
                        };
                        v.NumericColumnDecimals["Value"] = 2;
                    });

                    // Legend per domain
                    rs.SectionWithAnchor("Legend");
                    int lhdr = rs.CurrentRow;
                    rs.Sheet.Cell(lhdr, 1, "Status"); rs.Sheet.CellBold(lhdr, 1, true); rs.Sheet.CellBackground(lhdr, 1, "#F2F2F2");
                    rs.Sheet.Cell(lhdr, 2, "Meaning"); rs.Sheet.CellBold(lhdr, 2, true); rs.Sheet.CellBackground(lhdr, 2, "#F2F2F2");
                    var pal = StatusPalettes.Default;
                    rs.Sheet.Cell(lhdr + 1, 1, "OK"); if (pal.FillHexMap.TryGetValue("OK", out var ok2)) rs.Sheet.CellBackground(lhdr + 1, 1, ok2);
                    rs.Sheet.Cell(lhdr + 1, 2, "All checks passed or acceptable");
                    rs.Sheet.Cell(lhdr + 2, 1, "Warning"); if (pal.FillHexMap.TryGetValue("Warning", out var wr2)) rs.Sheet.CellBackground(lhdr + 2, 1, wr2);
                    rs.Sheet.Cell(lhdr + 2, 2, "Requires attention; not blocking");
                    rs.Sheet.Cell(lhdr + 3, 1, "Error"); if (pal.FillHexMap.TryGetValue("Error", out var er2)) rs.Sheet.CellBackground(lhdr + 3, 1, er2);
                    rs.Sheet.Cell(lhdr + 3, 2, "Blocking or invalid configuration");
                    rs.Spacer();

                    if (d.Recommendations.Length > 0) {
                        rs.SectionWithAnchor("Recommendations").BulletedList(d.Recommendations);
                    }
                    if (d.Positives.Length > 0) {
                        rs.SectionWithAnchor("Positives").BulletedList(d.Positives);
                    }
                    rs.SectionWithAnchor("References");
                    rs.References(d.References).Finish(autoFitColumns: false);
                }

                // TOC for easy navigation (first sheet)
                doc.AddTableOfContents(placeFirst: true, includeNamedRanges: false);

                // Validate OpenXML and print any issues (helps catch Excel repair causes)
                var issues = doc.ValidateOpenXml();
                if (issues.Count > 0) {
                    Console.WriteLine($"[!] OpenXML validation issues: {issues.Count}");
                    int show = Math.Min(issues.Count, 15);
                    for (int i = 0; i < show; i++) Console.WriteLine(" - " + issues[i]);
                } else {
                    Console.WriteLine("[+] OpenXML validation clean");
                }

                // Save (no Excel launch yet)
                doc.Save(false);

                // Re-open from disk and verify properties + header/footer
                using (var verify = ExcelDocument.Load(filePath, readOnly: true))
                {
                    Console.WriteLine("[=] Verifying saved workbook properties and header/footer...");
                    Console.WriteLine("    Title   : " + (verify.BuiltinDocumentProperties.Title ?? "<null>"));
                    Console.WriteLine("    Author  : " + (verify.BuiltinDocumentProperties.Creator ?? "<null>"));
                    Console.WriteLine("    Company : " + (verify.ApplicationProperties.Company ?? "<null>"));

                    var summary = verify.Sheets.FirstOrDefault(s => string.Equals(s.Name, "Summary", System.StringComparison.Ordinal));
                    if (summary != null)
                    {
                        var hf = summary.GetHeaderFooter();
                        Console.WriteLine("    Header (L/C/R): [" + hf.HeaderLeft + "] [" + hf.HeaderCenter + "] [" + hf.HeaderRight + "]");
                        Console.WriteLine("    Footer (L/C/R): [" + hf.FooterLeft + "] [" + hf.FooterCenter + "] [" + hf.FooterRight + "]");
                        Console.WriteLine("    Header has &G: " + hf.HeaderHasPicturePlaceholder + ", Footer has &G: " + hf.FooterHasPicturePlaceholder);
                    }
                    else
                    {
                        Console.WriteLine("    Summary sheet not found for header/footer verification.");
                    }
                }

                if (openExcel)
                {
                    // Open in Excel now, after verification
                    doc.Open(filePath, true);
                }
            }
        }
    }
}
