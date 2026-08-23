using System.Net;
using System.Text;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PdfHtmlScenarioBuilder {
    internal static string Create(PdfBenchmarkScenario scenario) {
        var html = new StringBuilder();
        html.AppendLine("<!doctype html><html lang=\"en-US\"><head><meta charset=\"utf-8\"><style>");
        html.AppendLine("@page { size: A4; margin: 12mm; }");
        html.AppendLine("body { font-family: sans-serif; font-size: 9pt; color: #172033; margin: 0; }");
        html.AppendLine(".report-page { break-after: page; page-break-after: always; }");
        html.AppendLine(".report-page:last-child { break-after: auto; page-break-after: auto; }");
        html.AppendLine("h1 { font-size: 18pt; margin: 0 0 8pt; } p { margin: 0 0 5pt; }");
        html.AppendLine("table { width: 100%; border-collapse: collapse; margin-top: 8pt; }");
        html.AppendLine("th, td { border: 0.5pt solid #7a8497; padding: 3pt; text-align: left; }");
        html.AppendLine("th { background: #dceffc; font-weight: 700; }</style></head><body>");
        for (int page = 1; page <= scenario.PageCount; page++) {
            html.AppendLine("<section class=\"report-page\">");
            html.Append("<h1>").Append(WebUtility.HtmlEncode(scenario.PageTitle(page))).AppendLine("</h1>");
            for (int paragraph = 0; paragraph < scenario.ParagraphsPerPage; paragraph++) {
                html.Append("<p>").Append(WebUtility.HtmlEncode(scenario.Narrative(page, paragraph))).AppendLine("</p>");
            }
            html.AppendLine("<table><thead><tr>");
            IReadOnlyList<string[]> rows = scenario.TableRows(page);
            foreach (string heading in rows[0]) {
                html.Append("<th>").Append(WebUtility.HtmlEncode(heading)).AppendLine("</th>");
            }
            html.AppendLine("</tr></thead><tbody>");
            foreach (string[] row in rows.Skip(1)) {
                html.AppendLine("<tr>");
                foreach (string cell in row) {
                    html.Append("<td>").Append(WebUtility.HtmlEncode(cell)).AppendLine("</td>");
                }
                html.AppendLine("</tr>");
            }
            html.AppendLine("</tbody></table></section>");
        }
        html.AppendLine("</body></html>");
        return html.ToString();
    }
}
