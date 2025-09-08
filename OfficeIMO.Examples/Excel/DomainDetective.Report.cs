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

                report.TableFrom(data, title: "Domains", configure: opts => {
                    opts.ExpandProperties.Add(nameof(MailDomainClassificationResult.ScoreBreakdown));
                    opts.HeaderCase = HeaderCase.Title;
                    opts.NullPolicy = NullPolicy.EmptyString;
                    opts.Ignore = new[] { nameof(MailDomainClassificationResult.Recommendations), nameof(MailDomainClassificationResult.Positives), nameof(MailDomainClassificationResult.References) };
                });

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
                    });

                    if (d.Recommendations.Length > 0) {
                        rs.Section("Recommendations").BulletedList(d.Recommendations);
                    }
                    if (d.Positives.Length > 0) {
                        rs.Section("Positives").BulletedList(d.Positives);
                    }
                    rs.References(d.References);
                }

                // Save
                doc.Save(openExcel);
            }
        }
    }
}

