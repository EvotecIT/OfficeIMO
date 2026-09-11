using System.Globalization;
using System.Net;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Project;

namespace OfficeIMO.Workflows;

public static partial class ProjectReportWorkflow {
    private static readonly Lazy<string> ProjectReportStyles = new(() => {
        using var stream = typeof(ProjectReportWorkflow).Assembly.GetManifestResourceStream("OfficeIMO.Workflows.Project.ProjectReport.css")
            ?? throw new InvalidOperationException("The embedded project report stylesheet is missing.");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    });
    /// <summary>Creates a self-contained responsive HTML report with printable vector pages and an accessible data table.</summary>
    public static string ToHtml(ProjectView view, OfficeRenderingProfile? typography = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(view);
        var body = new StringBuilder("<style>").Append(ProjectReportStyles.Value)
            .Append("@page{size:").Append(view.PageWidth.ToString(CultureInfo.InvariantCulture)).Append("pt ")
            .Append(view.PageHeight.ToString(CultureInfo.InvariantCulture)).Append("pt;margin:0}</style><main class=\"project-report\"><header>");
        body.Append("<h1>").Append(WebUtility.HtmlEncode(view.Title)).Append("</h1><p><a href=\"#report-data\">View report data</a> · Wide diagrams scroll horizontally on small screens.</p></header><div class=\"project-pages\" tabindex=\"0\" role=\"region\" aria-label=\"Report diagrams\">");
        int number = 0;
        foreach (string svg in ToSvg(view, typography, cancellationToken)) body.Append("<section class=\"project-page\" aria-label=\"Report page ").Append(++number).Append("\">").Append(svg).Append("</section>");
        body.Append("</div><div id=\"report-data\" class=\"project-data\" tabindex=\"0\" role=\"region\" aria-label=\"Report data\"><table><caption>Report data</caption><thead><tr>");
        foreach (var column in view.Columns) body.Append("<th scope=\"col\">").Append(WebUtility.HtmlEncode(ProjectView.ColumnTitle(column))).Append("</th>");
        body.Append("<th scope=\"col\">Status</th>");
        body.Append("</tr></thead><tbody>");
        foreach (var row in view.Rows) {
            cancellationToken.ThrowIfCancellationRequested(); body.Append("<tr>");
            foreach (var column in view.Columns) body.Append("<td>").Append(WebUtility.HtmlEncode(row.GetText(column))).Append("</td>");
            body.Append("<td>").Append(view.Kind == ProjectViewKind.ResourceUsage || view.Kind == ProjectViewKind.ResourceHistogram ? "Resource" : row.IsSummary ? "Summary" : "Task").Append(row.IsCritical ? "; critical" : "").Append("</td>");
            body.Append("</tr>");
        }
        body.Append("</tbody></table></div>");
        if (IsUsage(view)) {
            AppendHtmlTable(body, "Work hours", new[] { "UID", "Name" }.Concat(view.Buckets.Select(b => b.Start.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture))),
                view.Rows.Select(row => new[] { row.Uid.ToString(System.Globalization.CultureInfo.InvariantCulture), row.Name }.Concat(row.BucketWorkHours.Select(v => v.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)))), cancellationToken);
        }
        if (view.Links.Count > 0) AppendHtmlTable(body, "Dependencies", new[] { "Predecessor UID", "Successor UID", "Type", "Lag" },
            view.Links.Select(link => new[] { link.PredecessorUid.ToString(System.Globalization.CultureInfo.InvariantCulture), link.SuccessorUid.ToString(System.Globalization.CultureInfo.InvariantCulture), DependencyText(link.Type), link.LagText }), cancellationToken);
        if (view.Rows.Any(r => r.BaselineStart.HasValue || r.BaselineFinish.HasValue)) AppendHtmlTable(body, "Baseline dates", new[] { "UID", "Name", "Baseline start", "Baseline finish" },
            view.Rows.Select(row => new[] { row.Uid.ToString(System.Globalization.CultureInfo.InvariantCulture), row.Name, row.BaselineStart?.ToString("yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture) ?? "", row.BaselineFinish?.ToString("yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture) ?? "" }), cancellationToken);
        body.Append("</main>");
        return OfficeHtmlDocumentShell.WrapBody(body.ToString(), new OfficeHtmlDocumentOptions { Title = view.Title, IncludeDefaultStyles = false });
    }

    private static void AppendHtmlTable(StringBuilder body, string title, IEnumerable<string> headers, IEnumerable<IEnumerable<string>> rows, CancellationToken token) {
        body.Append("<div class=\"project-data\" tabindex=\"0\" role=\"region\" aria-label=\"").Append(WebUtility.HtmlEncode(title)).Append("\"><table><caption>").Append(WebUtility.HtmlEncode(title)).Append("</caption><thead><tr>");
        foreach (string header in headers) body.Append("<th scope=\"col\">").Append(WebUtility.HtmlEncode(header)).Append("</th>");
        body.Append("</tr></thead><tbody>");
        foreach (var row in rows) {
            token.ThrowIfCancellationRequested(); body.Append("<tr>");
            foreach (string value in row) body.Append("<td>").Append(WebUtility.HtmlEncode(value)).Append("</td>");
            body.Append("</tr>");
        }
        body.Append("</tbody></table></div>");
    }

}
