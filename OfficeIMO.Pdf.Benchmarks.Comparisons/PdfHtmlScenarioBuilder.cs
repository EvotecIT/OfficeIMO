using System.Net;
using System.Text;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PdfHtmlScenarioBuilder {
    internal static string Create(PdfBenchmarkScenario scenario) {
        var html = new StringBuilder();
        AppendLine(html, "<!doctype html><html lang=\"en-US\"><head><meta charset=\"utf-8\"><style>");
        AppendLine(html, "@page { size: A4; margin: 12mm; }");
        AppendLine(html, "body { font-family: sans-serif; font-size: 9pt; color: #172033; margin: 0; }");
        AppendLine(html, ".report-page { break-after: page; page-break-after: always; }");
        AppendLine(html, ".report-page:last-child { break-after: auto; page-break-after: auto; }");
        AppendLine(html, "h1 { font-size: 18pt; margin: 0 0 8pt; } p { margin: 0 0 5pt; }");
        AppendLine(html, "table { width: 100%; border-collapse: collapse; margin-top: 8pt; }");
        AppendLine(html, "th, td { border: 0.5pt solid #7a8497; padding: 3pt; text-align: left; }");
        AppendLine(html, "th { background: #dceffc; font-weight: 700; }</style></head><body>");
        for (int page = 1; page <= scenario.PageCount; page++) {
            AppendLine(html, "<section class=\"report-page\">");
            html.Append("<h1>").Append(WebUtility.HtmlEncode(scenario.PageTitle(page))).Append("</h1>\n");
            for (int paragraph = 0; paragraph < scenario.ParagraphsPerPage; paragraph++) {
                html.Append("<p>").Append(WebUtility.HtmlEncode(scenario.Narrative(page, paragraph))).Append("</p>\n");
            }
            AppendLine(html, "<table><thead><tr>");
            IReadOnlyList<string[]> rows = scenario.TableRows(page);
            foreach (string heading in rows[0]) {
                html.Append("<th>").Append(WebUtility.HtmlEncode(heading)).Append("</th>\n");
            }
            AppendLine(html, "</tr></thead><tbody>");
            foreach (string[] row in rows.Skip(1)) {
                AppendLine(html, "<tr>");
                foreach (string cell in row) {
                    html.Append("<td>").Append(WebUtility.HtmlEncode(cell)).Append("</td>\n");
                }
                AppendLine(html, "</tr>");
            }
            AppendLine(html, "</tbody></table></section>");
        }
        AppendLine(html, "</body></html>");
        return html.ToString();
    }

    private static void AppendLine(StringBuilder builder, string value) =>
        builder.Append(value).Append('\n');
}
