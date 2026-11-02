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
    /// Classic (baseline) Excel-only demo using simple ExcelReportBuilder blocks and direct cell writes
    /// without the newer convenience helpers. Intended as a reference for users who prefer
    /// explicit, standard building techniques.
    /// </summary>
    internal static class DomainDetectiveReportSheetsClassic
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
            Console.WriteLine("[*] Excel - Domain Detective Sheets (Classic) demo");
            string filePath = Path.Combine(folderPath, "DomainDetective.Report.Sheets.Classic.xlsx");
            var rows = GenerateFakeData(50, seed: 42);
            using var doc = ExcelDocument.Create(filePath);
            {
                doc.BuiltinDocumentProperties.Title = "Domain Detective — Classic Sheets";
                doc.BuiltinDocumentProperties.Creator = "OfficeIMO";
                doc.ApplicationProperties.Company = "Evotec";
                doc.ApplicationProperties.ApplicationName = "OfficeIMO.Excel";
                doc.BuiltinDocumentProperties.Keywords = "excel,report,sheets,classic";
            }

            // Overview
            var overview = doc.AddSheet("Overview", ExcelSheetNameValidationMode.Sanitize).Report();
            overview.Title("Domain Detective — Overview", $"Generated {DateTime.Now:yyyy-MM-dd HH:mm}");
            int totalWarnings = rows.Sum(x => x.WarningCount);
            int totalErrors = rows.Sum(x => x.ErrorCount);
            int atRisk = rows.Count(x => x.Status.Equals("Error", StringComparison.OrdinalIgnoreCase) || x.WarningCount > 0);
            // KPIs via basic PropertiesGrid (no KpiRow helper)
            overview.PropertiesGrid(new (string, object? )[] { ("Domains", rows.Count), ("At Risk", atRisk), ("Errors", totalErrors), ("Warnings", totalWarnings), ("Generated", DateTime.Now.ToString("yyyy-MM-dd")), ("Version", "v1") }, columns: 3);
            // Header/footer + logo (fixed URL)
            const string logoUrl = "https://evotec.pl/wp-content/uploads/2015/05/Logo-evotec-012.png";
            overview.Sheet.SetHeaderFooter(headerLeft: "Domain Detective — Overview", headerRight: "Page &P of &N");
            await overview.Sheet.SetHeaderImageFromUrlAsync(ExcelHeaderFooterPosition.Right, logoUrl, width: OfficeLength.Points(96), height: OfficeLength.Points(32));
            await overview.AddImageAsync(1, 6, new System.Uri(logoUrl), new OfficeIMO.Excel.ExcelImageOptions {
                Width = OfficeIMO.OfficeLength.Pixels(120),
                Height = OfficeIMO.OfficeLength.Pixels(40)
            });
            // Summary table
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
                opts.PinnedFirst = new[]
                {
                    nameof(DomainRow.Domain)
                };
            }, visuals: v =>
            {
                v.IconSetColumns.Add("Score");
                (v.TextBackgrounds["Status"], v.BoldByText["Status"]) = ReportStatusColors.Create(OfficeTheme.Default);
                v.AutoFormatDecimals = 2;
            }).Range!;
            // Links from Domain column to detail sheet
            overview.Sheet.LinkByHeaderToInternalSheetsInRange(summaryRange, header: "Domain", targetA1: "A1", styled: true);
            // Print settings (explicit, not via PrintDefaults)
            overview.Sheet.SetGridlinesVisible(false);
            overview.Sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0);
            doc.SetPrintArea(overview.Sheet, summaryRange);
            // Manual legend (no SectionLegend helper)
            overview.Section("Legend");
            int hdr = overview.CurrentRow;
            overview.Sheet.Cell(hdr, 1, "Status");
            overview.Sheet.CellBold(hdr, 1, true);
            overview.Sheet.CellBackground(hdr, 1, "#F2F2F2");
            overview.Sheet.Cell(hdr, 2, "Meaning");
            overview.Sheet.CellBold(hdr, 2, true);
            overview.Sheet.CellBackground(hdr, 2, "#F2F2F2");
            overview.Sheet.Cell(hdr, 3, "Recommended Action");
            overview.Sheet.CellBold(hdr, 3, true);
            overview.Sheet.CellBackground(hdr, 3, "#F2F2F2");
            var pal = ReportStatusColors.Create(OfficeTheme.Default);
            overview.Sheet.Cell(hdr + 1, 1, "OK");
            if (pal.FillHexMap.TryGetValue("OK", out var ok))
                overview.Sheet.CellBackground(hdr + 1, 1, ok);
            overview.Sheet.Cell(hdr + 1, 2, "All checks passed or acceptable");
            overview.Sheet.Cell(hdr + 1, 3, "No action required");
            overview.Sheet.Cell(hdr + 2, 1, "Warning");
            if (pal.FillHexMap.TryGetValue("Warning", out var wr))
                overview.Sheet.CellBackground(hdr + 2, 1, wr);
            overview.Sheet.Cell(hdr + 2, 2, "Requires attention; not blocking");
            overview.Sheet.Cell(hdr + 2, 3, "Review recommendations");
            overview.Sheet.Cell(hdr + 3, 1, "Error");
            if (pal.FillHexMap.TryGetValue("Error", out var er))
                overview.Sheet.CellBackground(hdr + 3, 1, er);
            overview.Sheet.Cell(hdr + 3, 2, "Blocking or invalid configuration");
            overview.Sheet.Cell(hdr + 3, 3, "Fix immediately");
            overview.Spacer(rows: 4);
            // Finish overview
            overview.Finish(autoFitColumns: false);
            DomainDetectiveReport.SizeSummaryColumns(overview.Sheet, summaryRange);

            // Detail sheets (classic)
            foreach (var d in rows)
                BuildDomainSheet(doc, d);
            // TOC + backlinks
            doc.AddTableOfContents(sheetName: "Index", placeFirst: true, includeNamedRanges: false);
            doc.AddBackLinksToToc(tocSheetName: "Index", row: 2, col: 1, text: "← Index");
            doc.Save();
            if (openExcel)
                doc.OpenInApplication(filePath);
        }

        private static void BuildDomainSheet(ExcelDocument doc, DomainRow d)
        {
            var s = doc.AddSheet(d.Domain, ExcelSheetNameValidationMode.Sanitize).Report();
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
            // Legend (manual)
            s.SectionWithAnchor("Legend");
            int lhdr = s.CurrentRow;
            s.Sheet.Cell(lhdr, 1, "Status");
            s.Sheet.CellBold(lhdr, 1, true);
            s.Sheet.CellBackground(lhdr, 1, "#F2F2F2");
            s.Sheet.Cell(lhdr, 2, "Meaning");
            s.Sheet.CellBold(lhdr, 2, true);
            s.Sheet.CellBackground(lhdr, 2, "#F2F2F2");
            s.Sheet.Cell(lhdr, 3, "Recommended Action");
            s.Sheet.CellBold(lhdr, 3, true);
            s.Sheet.CellBackground(lhdr, 3, "#F2F2F2");
            var pal = ReportStatusColors.Create(OfficeTheme.Default);
            if (pal.FillHexMap.TryGetValue("OK", out var ok2))
            {
                s.Sheet.Cell(lhdr + 1, 1, "OK");
                s.Sheet.CellBackground(lhdr + 1, 1, ok2);
                s.Sheet.Cell(lhdr + 1, 2, "All checks passed or acceptable");
                s.Sheet.Cell(lhdr + 1, 3, "No action required");
            }

            if (pal.FillHexMap.TryGetValue("Warning", out var wr2))
            {
                s.Sheet.Cell(lhdr + 2, 1, "Warning");
                s.Sheet.CellBackground(lhdr + 2, 1, wr2);
                s.Sheet.Cell(lhdr + 2, 2, "Requires attention; not blocking");
                s.Sheet.Cell(lhdr + 2, 3, "Review recommendations");
            }

            if (pal.FillHexMap.TryGetValue("Error", out var er2))
            {
                s.Sheet.Cell(lhdr + 3, 1, "Error");
                s.Sheet.CellBackground(lhdr + 3, 1, er2);
                s.Sheet.Cell(lhdr + 3, 2, "Blocking or invalid configuration");
                s.Sheet.Cell(lhdr + 3, 3, "Fix immediately");
            }

            s.Spacer(rows: 4);
            // Recommendations (manual bullet + fill)
            if (d.Recommendations.Length > 0)
            {
                s.SectionWithAnchor("Recommendations");
                int start = s.CurrentRow;
                s.BulletedList(d.Recommendations);
                int end = start + d.Recommendations.Length - 1;
                for (int r = start; r <= end; r++)
                    s.Sheet.CellBackground(r, 1, "#FFF4CE");
            }

            // Positives (manual bullet + fill)
            if (d.Positives.Length > 0)
            {
                s.SectionWithAnchor("Positives");
                int startP = s.CurrentRow;
                s.BulletedList(d.Positives);
                int endP = startP + d.Positives.Length - 1;
                for (int r = startP; r <= endP; r++)
                    s.Sheet.CellBackground(r, 1, "#E7F4E4");
            }

            if (d.References.Length > 0)
                s.SectionWithAnchor("References").References(d.References);
            s.Finish(autoFitColumns: false);
            // Label and value columns wide enough for their text, and each domain on one page.
            foreach ((int column, double width) in new[] { (1, 16d), (2, 30d), (3, 22d), (4, 20d), (5, 12d), (6, 10d) }) {
                s.Sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(width));
            }
            s.Sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 1);
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