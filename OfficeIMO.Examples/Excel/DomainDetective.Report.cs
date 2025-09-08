using System;
using System.Collections.Generic;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent.Report;
using OfficeIMO.Excel.Utilities;

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
                // Sheet 1: Summary of all domains as a table
                var report = new ReportSheetBuilder(doc, "Summary");
                report.Title("Domain Detective — Mail Classification Summary");

                var summaryRange = report.TableFrom(data, title: "Domains", configure: opts => {
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
                    // Also show explicit colors using SixLabors for convenience
                    viz.TextBackgrounds["Status"] = new System.Collections.Generic.Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
                    {
                        ["Error"] = "#F8C9C6",      // stronger light red
                        ["Warning"] = "#FFE59A",    // stronger light yellow
                        ["Success"] = "#CDEFCB",    // stronger light green
                        ["Ok"] = "#CDEFCB",
                        ["Pass"] = "#CDEFCB",
                    };
                    viz.BoldByText["Status"] = new HashSet<string>(StringComparer.OrdinalIgnoreCase){"Error","Warning"};
                });
                // Emphasize statuses (bold for Error/Warning)
                // Also request bolding via visuals fallback (works even when header row isn't the first row)
                // Note: we keep a direct call as a best-effort attempt; visuals will ensure it applies.
                try { report.Sheet.ColumnStyleByHeader("Status").BoldByTextSet(new HashSet<string>(StringComparer.OrdinalIgnoreCase){"Error","Warning"}); } catch { }

                // Make Summary sheet print nicely
                report.Sheet.SetGridlinesVisible(false);
                report.Sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0);
                doc.SetPrintArea(report.Sheet, summaryRange);
                report.Finish(autoFitColumns: true);

                // One sheet per domain with richer layout
                foreach (var d in data) {
                    var rs = new ReportSheetBuilder(doc, d.Domain);
                    rs.Title($"Mail Classification — {d.Domain}", d.Summary)
                      .Section("Overview")
                      .PropertiesGrid(new (string, object?)[] {
                          ("Domain", d.Domain),
                          ("Classification", d.Classification),
                          ("Confidence", d.Confidence),
                          ("Status", d.Status),
                          ("Warnings", d.WarningCount),
                          ("Errors", d.ErrorCount)
                      }, columns: 3)
                      .Section("Signals")
                      .PropertiesGrid(new (string, object?)[] {
                          ("Receiving", string.Join(", ", d.ReceivingSignals)),
                          ("Sending", string.Join(", ", d.SendingSignals))
                      }, columns: 2)
                      .Score("Score", d.Score)
                      .Section("Score Breakdown");

                    rs.TableFrom(d.ScoreBreakdown, title: null, configure: o => {
                        o.Columns = new[] { nameof(ScorePair.Name), nameof(ScorePair.Value) };
                        o.HeaderCase = HeaderCase.Title;
                    }, visuals: v => {
                        // Example of explicit icon thresholds for demo purposes
                        v.IconSets["Value"] = new OfficeIMO.Excel.Fluent.Report.IconSetOptions {
                            IconSet = DocumentFormat.OpenXml.Spreadsheet.IconSetValues.ThreeSymbols,
                            ShowValue = true,
                            ReverseOrder = false,
                            PercentThresholds = new double[] { 0, 60, 85 } // 0-60 red, 60-85 yellow, 85-100 green
                        };
                        v.NumericColumnDecimals["Value"] = 2;
                    });

                    if (d.Recommendations.Length > 0) {
                        rs.Section("Recommendations").BulletedList(d.Recommendations);
                    }
                    if (d.Positives.Length > 0) {
                        rs.Section("Positives").BulletedList(d.Positives);
                    }
                    rs.References(d.References).Finish(autoFitColumns: true);
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

                // Save
                doc.Save(openExcel);
            }
        }
    }
}
