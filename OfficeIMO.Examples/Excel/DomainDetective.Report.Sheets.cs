using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using System.Threading.Tasks;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Side-by-side Excel-only demo using the "SheetXXX + Blocks" approach (backed by SheetComposer)
    /// to compare with the existing DomainDetective.Report.cs example. Generates ~50 domains.
    /// </summary>
    internal static class DomainDetectiveReportSheets {
        private record ScorePair(string Name, double Value);
        private record DomainRow(
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
            Console.WriteLine("[*] Excel - Domain Detective Sheets (Excelish) demo");
            string filePath = Path.Combine(folderPath, "DomainDetective.Report.Sheets.xlsx");
            var rows = GenerateFakeData(50, seed: 42);

            using var doc = ExcelDocument.Create(filePath);

            // Document properties
            doc.AsFluent().Info(i => i
                .Title("Domain Detective — Excelish Sheets")
                .Author("OfficeIMO")
                .Company("Evotec")
                .Application("OfficeIMO.Excel")
                .Keywords("excel,report,sheets,domains")
            ).End();

            // Overview sheet
            var overview = new SheetComposer(doc, "Overview");
            overview.Title("Domain Detective — Overview", $"Generated {DateTime.Now:yyyy-MM-dd HH:mm}");
            // Header/footer with Evotec logo + page text (fixed URL)
            const string logoUrl = "https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png";
            overview.HeaderLogoUrl(logoUrl, OfficeIMO.Excel.HeaderFooterPosition.Center, 120, 40, leftText: "Page &P of &N");
            // Also place the logo inside the sheet (first page) via URL
            overview.ImageFromUrlAt(row: 1, column: 6, url: logoUrl, widthPixels: 120, heightPixels: 40);

            int totalWarnings = rows.Sum(x => x.WarningCount);
            int totalErrors = rows.Sum(x => x.ErrorCount);
            int atRisk = rows.Count(x => x.Status.Equals("Error", StringComparison.OrdinalIgnoreCase) || x.WarningCount > 0);
            int okCount = rows.Count(x => string.Equals(x.Status, "OK", StringComparison.OrdinalIgnoreCase));
            int warnCount = rows.Count(x => string.Equals(x.Status, "Warning", StringComparison.OrdinalIgnoreCase));
            int errCount = rows.Count(x => string.Equals(x.Status, "Error", StringComparison.OrdinalIgnoreCase));

            // KPIs as compact cards (labels above values)
            overview.KpiRow(new (string, object?)[] {
                ("Domains", rows.Count),
                ("At Risk", atRisk),
                ("Errors", totalErrors),
                ("Warnings", totalWarnings),
                ("Generated", DateTime.Now.ToString("yyyy-MM-dd")),
                ("Version", "v1")
            }, perRow: 3);

            // At-a-glance columns (Excelish side-by-side blocks)
            overview.Section("At a Glance");
            overview.Columns(3, cols => {
                cols[0].Section("Totals").KeyValues(new (string, object?)[] {
                    ("Domains", rows.Count),
                    ("At Risk", atRisk),
                    ("Warnings", totalWarnings),
                    ("Errors", totalErrors)
                });
                cols[1].Section("Status Breakdown").KeyValues(new (string, object?)[] {
                    ("OK", okCount),
                    ("Warning", warnCount),
                    ("Error", errCount)
                });
                cols[2].Section("Tips").BulletedList(new[]{
                    "Use filters in the header row.",
                    "Click a Domain to open details.",
                    "Use ↑ Top links to navigate."
                });
            });

            // Summary table with dynamic ScoreBreakdown columns
            var summaryRange = overview.TableFrom(rows, title: "Domains", configure: opts => {
                opts.CollectionMapColumns[nameof(DomainRow.ScoreBreakdown)] = new CollectionColumnMapping {
                    KeyProperty = nameof(ScorePair.Name),
                    ValueProperty = nameof(ScorePair.Value)
                };
                opts.HeaderCase = HeaderCase.Title;
                opts.HeaderPrefixTrimPaths = new[] { nameof(DomainRow.ScoreBreakdown) + "." };
                opts.NullPolicy = NullPolicy.EmptyString;
                // Keep "Domain" as first column for easy linking
                opts.PinnedFirst = new[] { nameof(DomainRow.Domain) };
            }, visuals: v => {
                // Icon set for overall score
                v.IconSetColumns.Add("Score");
                // Emphasize status via background and bold
                (v.TextBackgrounds["Status"], v.BoldByText["Status"]) = StatusPalettes.Default;
                v.AutoFormatDecimals = 2;
            });

            // Link only the Domain column in the summary table to its detail sheet (styled)
            overview.Sheet.LinkByHeaderToInternalSheetsInRange(
                rangeA1: summaryRange,
                header: "Domain",
                targetA1: "A1",
                styled: true);

            // Make summary presentable for printing
            overview
                .PrintDefaults(showGridlines: false, fitToWidth: 1, fitToHeight: 0, printAreaA1: summaryRange)
                .Orientation(ExcelPageOrientation.Landscape)
                .Margins(ExcelMarginPreset.Narrow)
                .RepeatHeaderRows(1, 1);

            // Pretty legend (Status | Meaning | Recommended Action) using reusable SectionLegend
            {
                var pal = StatusPalettes.Default;
                var map = pal.FillHexMap; // case-insensitive
                overview.SectionLegend(
                    title: "Legend",
                    headers: new[] { "Status", "Meaning", "Recommended Action" },
                    rows: new[] {
                        new [] { "OK", "All checks passed or acceptable", "No action required" },
                        new [] { "Warning", "Requires attention; not blocking", "Review recommendations" },
                        new [] { "Error", "Blocking or invalid configuration", "Fix immediately" },
                    },
                    firstColumnFillByValue: map
                );
            }

            // Finish overview
            overview.Finish(autoFitColumns: true);

            // One detail sheet per domain (sequential by default)
            foreach (var d in rows) BuildDomainSheet(doc, d);

            // Build Index/TOC last so it includes every sheet, then add back-links
            SheetIndex.Add(doc, sheetName: "Index", placeFirst: true, includeNamedRanges: false);
            SheetIndex.AddBackLinks(doc, tocSheetName: "Index", row: 2, col: 1, text: "← Index");

            // Add header logo to Index (Left) with page number on Right for variety
            var idx = doc["Index"]; if (idx != null) {
                idx.HeaderFooter(h => {
                    h.Right("Page &P of &N");
                    h.LeftImageUrl(logoUrl, widthPoints: 96, heightPoints: 32);
                });
            }

            var errors = doc.ValidateOpenXml();
            if (errors.Count > 0) Console.WriteLine($"[!] OpenXML validation issues: {errors.Count}");

            doc.Save(false);
            if (openExcel) doc.Open(filePath, true);
        }

        private static void BuildDomainSheet(ExcelDocument doc, DomainRow d) {
            var s = new SheetComposer(doc, d.Domain);
            s.Title($"Mail Classification — {d.Domain}", d.Summary)
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

            s.TableFrom(d.ScoreBreakdown, title: null, configure: o => {
                o.Columns = new[] { nameof(ScorePair.Name), nameof(ScorePair.Value) };
                o.HeaderCase = HeaderCase.Title;
            }, visuals: v => {
                v.NumericColumnDecimals["Value"] = 2;
                v.IconSets["Value"] = new IconSetOptions {
                    IconSet = DocumentFormat.OpenXml.Spreadsheet.IconSetValues.ThreeSymbols,
                    ShowValue = true,
                    ReverseOrder = false,
                    PercentThresholds = new double[] { 0, 60, 85 }
                };
            });
            // Header logo on the Right, page number on Left (complements Index/Overview)
            s.HeaderLogoUrl("https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png",
                            OfficeIMO.Excel.HeaderFooterPosition.Right, 96, 32, leftText: "Page &P of &N");
            // Optional: embed the Evotec logo on each detail sheet near the title
            s.ImageFromUrlAt(row: 1, column: 5, url: "https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png", widthPixels: 100, heightPixels: 34);

            // Legend per domain (reusable SectionLegend)
            {
                var pal = StatusPalettes.Default;
                s.SectionLegend(
                    title: "Legend",
                    headers: new[] { "Status", "Meaning", "Recommended Action" },
                    rows: new[] {
                        new [] { "OK", "All checks passed or acceptable", "No action required" },
                        new [] { "Warning", "Requires attention; not blocking", "Review recommendations" },
                        new [] { "Error", "Blocking or invalid configuration", "Fix immediately" },
                    },
                    firstColumnFillByValue: pal.FillHexMap
                );
            }

            // Recommendations highlighted (warning tone) using BulletedListWithFill
            if (d.Recommendations.Length > 0) {
                s.SectionWithAnchor("Recommendations");
                s.BulletedListWithFill(d.Recommendations, fillHex: "#FFF4CE");
            }
            // Positives highlighted (success tone) using BulletedListWithFill
            if (d.Positives.Length > 0) {
                s.SectionWithAnchor("Positives");
                s.BulletedListWithFill(d.Positives, fillHex: "#E7F4E4");
            }
            if (d.References.Length > 0) s.SectionWithAnchor("References").References(d.References);

            s.Finish(autoFitColumns: true);
        }



        private static List<DomainRow> GenerateFakeData(int count, int seed) {
            var rnd = new Random(seed);
            string[] classes = new[] { "Sending", "Receiving", "SendingAndReceiving" };
            string[] confidences = new[] { "Low", "Medium", "High" };
            string[] recv = new[] { "MX", "TLS-RPT", "NullMX" };
            string[] send = new[] { "SPF", "DKIM", "DMARC", "BIMI" };

            var list = new List<DomainRow>(count);
            for (int i = 1; i <= count; i++) {
                string domain = $"domain-{i:000}.example";
                string cls = classes[rnd.Next(classes.Length)];
                string conf = confidences[rnd.Next(confidences.Length)];
                var recvS = recv.Where(_ => rnd.NextDouble() < 0.7).DefaultIfEmpty("MX").Distinct().ToArray();
                var sendS = send.Where(_ => rnd.NextDouble() < 0.7).DefaultIfEmpty("SPF").Distinct().ToArray();
                int warnings = rnd.Next(0, 7);
                int errors = rnd.NextDouble() < 0.15 ? rnd.Next(1, 3) : 0;
                string status = errors > 0 ? "Error" : (warnings > 0 ? "Warning" : "OK");
                int score = Math.Max(0, 10 - warnings - errors * 3);
                var breakdown = new List<ScorePair> {
                    new("HasMX", recvS.Contains("MX") ? 2 : 0),
                    new("HasNullMX", recvS.Contains("NullMX") ? -1 : 0),
                    new("EffectiveSPFSends", sendS.Contains("SPF") ? 2 : 0),
                    new("HasDKIM", sendS.Contains("DKIM") ? 2 : 0),
                    new("HasDMARC", sendS.Contains("DMARC") ? 2 : 0),
                    new("HasBIMI", sendS.Contains("BIMI") ? 1 : 0)
                };
                string summary = $"{cls} ({conf}); recv {recvS.Length}; send {sendS.Length}";
                string[] recs = warnings > 0 || !sendS.Contains("DMARC")
                    ? new[] { "Enable DMARC enforcement", "Rotate DKIM keys", "Review SPF flattening" }
                    : Array.Empty<string>();
                string[] pos = new[] { "SPF present", "DKIM present" }.Where(_ => rnd.NextDouble() < 0.8).ToArray();
                string[] refs = new[] {
                    "https://datatracker.ietf.org/doc/html/rfc7208",
                    "https://datatracker.ietf.org/doc/html/rfc6376",
                    "https://datatracker.ietf.org/doc/html/rfc7489"
                };

                list.Add(new DomainRow(domain, cls, conf, recvS, sendS, score, breakdown, status, warnings, errors, summary, recs, pos, refs));
            }
            return list;
        }
    }
}
