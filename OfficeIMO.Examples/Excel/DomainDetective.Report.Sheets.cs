using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using OfficeIMO.Data;
using System.Threading.Tasks;

namespace OfficeIMO.Examples.Excel
{
    /// <summary>
    /// Side-by-side Excel-only demo using the "SheetXXX + Blocks" approach (backed by ExcelReportBuilder)
    /// to compare with the existing DomainDetective.Report.cs example. Generates ~50 domains.
    /// </summary>
    internal static class DomainDetectiveReportSheets
    {
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
        public static async Task Example(string folderPath, bool openExcel)
        {
            Console.WriteLine("[*] Excel - Domain Detective Sheets (Excelish) demo");
            string filePath = Path.Combine(folderPath, "DomainDetective.Report.Sheets.xlsx");
            var rows = GenerateFakeData(50, seed: 42);
            using var doc = ExcelDocument.Create(filePath);
            // Document properties
            {
                doc.BuiltinDocumentProperties.Title = "Domain Detective — Excelish Sheets";
                doc.BuiltinDocumentProperties.Creator = "OfficeIMO";
                doc.ApplicationProperties.Company = "Evotec";
                doc.ApplicationProperties.ApplicationName = "OfficeIMO.Excel";
                doc.BuiltinDocumentProperties.Keywords = "excel,report,sheets,domains";
            }

            // Overview sheet
            var overview = doc.AddWorksheet("Overview", ExcelSheetNameValidationMode.Sanitize).Report();
            overview.Title("Domain Detective — Overview", $"Generated {DateTime.Now:yyyy-MM-dd HH:mm}");
            // Header/footer with Evotec logo + page text (fixed URL)
            const string logoUrl = "https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png";
            await overview.HeaderLogoFromUrlAsync(logoUrl, OfficeIMO.Excel.ExcelHeaderFooterPosition.Center, 120, 40, leftText: "Page &P of &N");
            // Also place the logo inside the sheet (first page) via URL
            await overview.AddImageFromUrlAsync(row: 1, column: 6, url: logoUrl, width: OfficeIMO.OfficeLength.Pixels(120), height: OfficeIMO.OfficeLength.Pixels(40));
            int totalWarnings = rows.Sum(x => x.WarningCount);
            int totalErrors = rows.Sum(x => x.ErrorCount);
            int atRisk = rows.Count(x => x.Status.Equals("Error", StringComparison.OrdinalIgnoreCase) || x.WarningCount > 0);
            int okCount = rows.Count(x => string.Equals(x.Status, "OK", StringComparison.OrdinalIgnoreCase));
            int warnCount = rows.Count(x => string.Equals(x.Status, "Warning", StringComparison.OrdinalIgnoreCase));
            int errCount = rows.Count(x => string.Equals(x.Status, "Error", StringComparison.OrdinalIgnoreCase));
            // KPIs as compact cards (labels above values)
            overview.KpiRow(new (string, object? )[] { ("Domains", rows.Count), ("At Risk", atRisk), ("Errors", totalErrors), ("Warnings", totalWarnings), ("Generated", DateTime.Now.ToString("yyyy-MM-dd")), ("Version", "v1") }, perRow: 3);
            // At-a-glance columns (Excelish side-by-side blocks)
            overview.Section("At a Glance");
            overview.Columns(3, cols =>
            {
                cols[0].Section("Totals").KeyValues(new (string, object? )[] { ("Domains", rows.Count), ("At Risk", atRisk), ("Warnings", totalWarnings), ("Errors", totalErrors) });
                cols[1].Section("Status Breakdown").KeyValues(new (string, object? )[] { ("OK", okCount), ("Warning", warnCount), ("Error", errCount) });
                cols[2].Section("Tips").BulletedList(new[] { "Use filters in the header row.", "Click a Domain to open details.", "Use ↑ Top links to navigate." });
            });
            // Summary table with dynamic ScoreBreakdown columns
            var summaryRange = overview.AddTable(rows, title: "Domains", configure: opts =>
            {
                opts.CollectionMapColumns[nameof(DomainRow.ScoreBreakdown)] = new CollectionColumnMapping
                {
                    KeyProperty = nameof(ScorePair.Name),
                    ValueProperty = nameof(ScorePair.Value)
                };
                opts.HeaderCase = HeaderCase.Title;
                opts.HeaderPrefixTrimPaths = new[]
                {
                    nameof(DomainRow.ScoreBreakdown) + "."
                };
                opts.NullPolicy = NullPolicy.EmptyString;
                // Keep "Domain" first, then Status/WarningCount/ErrorCount/Score, and push References to the end
                // Using the single-call convenience Order(pinFirst, priority, pinLast)
                opts.Order(pinFirst: new[] { nameof(DomainRow.Domain) }, priority: new[] { nameof(DomainRow.Status), nameof(DomainRow.WarningCount), nameof(DomainRow.ErrorCount), nameof(DomainRow.Score) }, pinLast: new[] { nameof(DomainRow.References) });
            }, visuals: v =>
            {
                // Icon set for overall score
                v.IconSetColumns.Add("Score");
                // Emphasize status via background and bold
                (v.TextBackgrounds["Status"], v.BoldByText["Status"]) = ReportStatusColors.Create(OfficeTheme.Default);
                v.AutoFormatDecimals = 2;
            }).Range!;
            // Link only the Domain column in the summary table to its detail sheet (styled)
            // Use non-throwing variant to avoid exceptions if headers change
            overview.Sheet.TryLinkByHeaderToInternalSheetsInRange(rangeA1: summaryRange, header: "Domain", targetA1: "A1", styled: true);
            // Make summary presentable for printing
            overview.PrintDefaults(showGridlines: false, fitToWidth: 1, fitToHeight: 0, printAreaA1: summaryRange).Orientation(OfficePageOrientation.Landscape).Margins(ExcelMarginPreset.Narrow).RepeatHeaderRows(1, 1);
            // Pretty legend (Status | Meaning | Recommended Action) using reusable SectionLegend
            {
                var pal = ReportStatusColors.Create(OfficeTheme.Default);
                var map = pal.FillHexMap.ToDictionary(pair => pair.Key, pair => OfficeColor.Parse(pair.Value), StringComparer.OrdinalIgnoreCase);
                overview.SectionLegend(title: "Legend", headers: new[] { "Status", "Meaning", "Recommended Action" }, rows: new[] { new[] { "OK", "All checks passed or acceptable", "No action required" }, new[] { "Warning", "Requires attention; not blocking", "Review recommendations" }, new[] { "Error", "Blocking or invalid configuration", "Fix immediately" }, }, firstColumnFillByValue: map);
            }

            // Finish overview
            // Overview has multiple side-by-side sections; use explicit column sizing per table
            overview.Finish(autoFitColumns: false);
            // One detail sheet per domain (sequential by default)
            foreach (var d in rows)
                await BuildDomainSheet(doc, d);
            // Build Index/TOC last so it includes every sheet, then add back-links
            doc.AddTableOfContents(sheetName: "Index", placeFirst: true, includeNamedRanges: false);
            doc.AddBackLinksToToc(tocSheetName: "Index", row: 2, col: 1, text: "← Index");
            // Add header logo to Index (Left) with page number on Right for variety
            var idx = doc["Index"];
            if (idx != null)
            {
                idx.SetHeaderFooter(headerRight: "Page &P of &N");
                await idx.SetHeaderImageFromUrlAsync(ExcelHeaderFooterPosition.Left, logoUrl, widthPoints: 96, heightPoints: 32);
            }

            var errors = doc.ValidateOpenXml();
            if (errors.Count > 0)
                Console.WriteLine($"[!] OpenXML validation issues: {errors.Count}");
            doc.Save();
            if (openExcel)
                doc.OpenInApplication(filePath);
        }

        private static async Task BuildDomainSheet(ExcelDocument doc, DomainRow d)
        {
            var s = doc.AddWorksheet(d.Domain, ExcelSheetNameValidationMode.Sanitize).Report();
            s.Title($"Mail Classification — {d.Domain}", d.Summary).Callout(d.ErrorCount > 0 ? OfficeCalloutKind.Danger : (string.Equals(d.Status, "Warning", StringComparison.OrdinalIgnoreCase) ? OfficeCalloutKind.Warning : OfficeCalloutKind.Info), "Status", $"Status: {d.Status}; Findings: {d.WarningCount} warning(s), {d.ErrorCount} error(s).").SectionWithAnchor("Overview").PropertiesGrid(new (string, object? )[] { ("Domain", d.Domain), ("Classification", d.Classification), ("Confidence", d.Confidence), ("Status", d.Status), ("Warnings", d.WarningCount), ("Errors", d.ErrorCount) }, columns: 3).SectionWithAnchor("Signals").PropertiesGrid(new (string, object? )[] { ("Receiving", string.Join(", ", d.ReceivingSignals)), ("Sending", string.Join(", ", d.SendingSignals)) }, columns: 2).Score("Score", d.Score).SectionWithAnchor("Score Breakdown");
            s.AddTable(d.ScoreBreakdown, title: null, configure: o =>
            {
                o.Columns = new[]
                {
                    nameof(ScorePair.Name),
                    nameof(ScorePair.Value)
                };
                o.HeaderCase = HeaderCase.Title;
            }, visuals: v =>
            {
                v.NumericColumnDecimals["Value"] = 2;
                v.IconSets["Value"] = new ExcelIconSetOptions
                {
                    IconSet = ExcelIconSet.ThreeSymbols,
                    ShowValue = true,
                    ReverseOrder = false,
                    PercentThresholds = new double[]
                    {
                        0,
                        60,
                        85
                    }
                };
            });
            // Header logo on the Right, page number on Left (complements Index/Overview)
            await s.HeaderLogoFromUrlAsync("https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png", OfficeIMO.Excel.ExcelHeaderFooterPosition.Right, 96, 32, leftText: "Page &P of &N");
            // Optional: embed the Evotec logo on each detail sheet near the title
            await s.AddImageFromUrlAsync(row: 1, column: 5, url: "https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png", width: OfficeIMO.OfficeLength.Pixels(100), height: OfficeIMO.OfficeLength.Pixels(34));
            // Legend per domain (reusable SectionLegend)
            {
                var pal = ReportStatusColors.Create(OfficeTheme.Default);
                s.SectionLegend(title: "Legend", headers: new[] { "Status", "Meaning", "Recommended Action" }, rows: new[] { new[] { "OK", "All checks passed or acceptable", "No action required" }, new[] { "Warning", "Requires attention; not blocking", "Review recommendations" }, new[] { "Error", "Blocking or invalid configuration", "Fix immediately" }, }, firstColumnFillByValue: pal.FillHexMap.ToDictionary(pair => pair.Key, pair => OfficeColor.Parse(pair.Value), StringComparer.OrdinalIgnoreCase));
            }

            // Recommendations highlighted (warning tone) using BulletedListWithFill
            if (d.Recommendations.Length > 0)
            {
                s.SectionWithAnchor("Recommendations");
                s.BulletedListWithFill(d.Recommendations, fill: OfficeColor.Parse("#FFF4CE"));
            }

            // Positives highlighted (success tone) using BulletedListWithFill
            if (d.Positives.Length > 0)
            {
                s.SectionWithAnchor("Positives");
                s.BulletedListWithFill(d.Positives, fill: OfficeColor.Parse("#E7F4E4"));
            }

            if (d.References.Length > 0)
                s.SectionWithAnchor("References").References(d.References);
            s.Finish(autoFitColumns: false);
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