using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;

/// <summary>Independent value, relationship and page-fragment checks for portable cell links.</summary>
internal static class PortableLinkVerifier {
    private const string Target = "https://example.com/report?q=1&other=2#details";
    private const string Tooltip = "Report _x0041_ <value> 🧪";

    internal static object Verify(string path) {
        if (path.EndsWith(".xlsx", StringComparison.Ordinal)) {
            WorkbookVerifier.Verify(path);
            using var sdk = SpreadsheetDocument.Open(path, false);
            var part = sdk.WorkbookPart!.WorksheetParts.Single();
            var links = part.Worksheet!.Descendants<Hyperlink>().ToArray();
            WorkbookVerifier.Require(links.Select(l => l.Reference!.Value).SequenceEqual(new[] { "A4", "B4", "A5" }), "Portable XLSX link coordinates differ.");
            WorkbookVerifier.Require(part.HyperlinkRelationships.Single(r => r.Id == links[0].Id!.Value).Uri.AbsoluteUri == Target &&
                part.HyperlinkRelationships.Single(r => r.Id == links[1].Id!.Value).Uri.AbsoluteUri == "mailto:report@example.com", "Portable XLSX targets differ.");
            WorkbookVerifier.Require(links[0].Tooltip!.Value == "Report _x005F_x0041_ <value> 🧪", "Literal XLSX tooltip escape differs.");
            using var reader = ExcelDocument.OpenDataReader(path, new ExcelReadOptions { A1Range = "A4:B4", HasHeaderRow = false });
            WorkbookVerifier.Require(reader.Read() && Convert.ToDouble(reader.GetValue(0)) == 12.5 &&
                reader.GetValue(1) is DateTime date && date == new DateTime(2026, 10, 8), "Linked XLSX numeric/date values differ.");
            return new { file = Path.GetFileName(path), links = links.Length, readers = 2, openXml = true };
        }
        var pdf = PdfReadDocument.Open(path);
        WorkbookVerifier.Require(!pdf.RepairReport.HasRepairs, "Portable PDF requires reader repairs.");
        bool spans = Path.GetFileName(path).Contains("spans", StringComparison.Ordinal);
        int count = 0;
        foreach (var page in pdf.Pages) {
            var links = page.GetLinkAnnotations(); count += links.Count;
            WorkbookVerifier.Require(links.Count >= 1, "PDF page lost a linked cell fragment.");
            foreach (var link in links) {
                WorkbookVerifier.Require(link.Uri == Target || spans && link.Uri == "mailto:report@example.com", "PDF URI action differs.");
                WorkbookVerifier.Require(link.Uri != Target || link.Contents == Tooltip, "PDF Unicode tooltip differs.");
                WorkbookVerifier.Require(link.Width > 0 && link.Height > 0 && link.Y1 >= 36 && link.Y2 <= 559.277, "PDF link rectangle is outside the table's page bounds.");
            }
        }
        WorkbookVerifier.Require(pdf.Pages.Count > (spans ? 1 : 3), "Portable PDF continuation fixture is incomplete.");
        if (!spans) {
            string body = string.Concat(pdf.Pages.SelectMany(p => p.GetTextSpans()).Select(s => s.Text).Where(t => t != "Value" && !t.StartsWith("Page ", StringComparison.Ordinal) && !int.TryParse(t, out _)));
            WorkbookVerifier.Require(body == string.Concat(Enumerable.Repeat("ABCDEFGHIJKLMNOPQRSTUVWXYZ", 300)), "Linked PDF continuation lost or duplicated text.");
        }
        return new { file = Path.GetFileName(path), pages = pdf.Pages.Count, links = count, repairs = 0 };
    }
}
