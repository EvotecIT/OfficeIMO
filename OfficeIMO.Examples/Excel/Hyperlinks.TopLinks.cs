using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates clickable hyperlinks in Excel tables using the Fluent API + RowEdit.SetFormula.
    /// Produces a small sheet with a "Top Links" table and converts the Title column to =HYPERLINK(Url, "Title").
    /// </summary>
    internal static class HyperlinksTopLinks {
        private sealed record LinkItem(string Title, string Url);

        public static void Example(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Hyperlinks in table cells (Top Links) example");
            string filePath = Path.Combine(folderPath, "Excel.Hyperlinks.TopLinks.xlsx");

            // Sample links (mix of RFC and vendor docs). Titles are human-friendly; Url is the canonical link.
            var links = new List<LinkItem> {
                new("RFC 7208 — SPF: Sender Policy Framework", "https://www.rfc-editor.org/rfc/rfc7208"),
                new("RFC 6376 — DKIM: DomainKeys Identified Mail", "https://www.rfc-editor.org/rfc/rfc6376"),
                new("RFC 7489 — DMARC: Authentication, Reporting, Conformance", "https://www.rfc-editor.org/rfc/rfc7489"),
                new("Microsoft Docs — DMARC Configure", "https://learn.microsoft.com/en-us/defender-office-365/email-authentication-dmarc-configure"),
                new("Microsoft Docs — SPF Configure", "https://learn.microsoft.com/en-us/defender-office-365/email-authentication-spf-configure"),
            };

            using var doc = ExcelDocument.Create(filePath);
            doc.AsFluent().Info(i => i
                .Title("OfficeIMO — Clickable Hyperlinks Example")
                .Author("OfficeIMO")
                .Company("Evotec")
                .Application("OfficeIMO.Excel")
            ).End();

            var s = new SheetComposer(doc, "Top Links");
            s.Title("Top Links", "Demonstrates =HYPERLINK(Url, Title) in a table");

            // Build a simple table: Title | Url
            var rows = links.Select(l => new { l.Title, l.Url }).ToList();
            var a1 = s.TableFrom(rows, title: null, configure: o => { o.HeaderCase = HeaderCase.Title; }, visuals: v => v.FreezeHeaderRow = true);

            // Convert Title cells to clickable hyperlinks using formulas (keeps Url visible for auditing)
            try {
                foreach (var row in s.Sheet.RowsObjects(a1)) {
                    var titleCell = row.CellByHeader("Title");
                    var urlCell = row.CellByHeader("Url");
                    string urlRef = IndexToCol(urlCell.ColumnIndex) + titleCell.RowIndex.ToString();
                    string safeTitle = (row.GetOrDefault<string>("Title", string.Empty) ?? string.Empty).Replace("\"", "\"\"");
                    row.SetFormula("Title", $"=HYPERLINK({urlRef},\"{safeTitle}\")");
                }
            } catch { }

            s.Finish(autoFitColumns: true);

            doc.Save(false);
            if (openExcel) doc.Open(filePath, true);
        }

        // Minimal helper to convert column index to Excel letter (duplicated here for a self-contained example)
        private static string IndexToCol(int index) {
            int dividend = index; string col = string.Empty;
            while (dividend > 0) { int modulo = (dividend - 1) % 26; col = Convert.ToChar('A' + modulo) + col; dividend = (dividend - modulo) / 26; }
            return col;
        }
    }
}

