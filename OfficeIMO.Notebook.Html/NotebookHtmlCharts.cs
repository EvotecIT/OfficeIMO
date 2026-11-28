using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using OfficeIMO.Markdown;
using OfficeIMO.Notebook.Markdown;

namespace OfficeIMO.Notebook.Html;

/// <summary>Writes an offline chart viewport over the native vector drawing and its saved data locations.</summary>
internal sealed class NotebookHtmlCharts(CancellationToken cancellationToken) {
    private const int MaxViews = 128;
    private const int MaxMarksPerView = 8192;
    private const int MaxLabelCharacters = 400;
    private const int MaxCharacters = 24 * 1024 * 1024;
    private static readonly Lazy<string> Assets = new(() =>
        "<style>\n" + ReadAsset("css") + "\n</style><script>\n" + ReadAsset("js") + "\n</script>");

    private readonly List<OfficeDiagnostic> _diagnostics = new();
    private int _views;
    private int _characters;

    internal IReadOnlyList<OfficeDiagnostic> Diagnostics => _diagnostics;

    internal IMarkdownBlock? Render(NotebookChartPreview chart, string location) {
        cancellationToken.ThrowIfCancellationRequested();
        int remaining = MaxCharacters - _characters;
        if (_views >= MaxViews || chart.ImageUri.Length >= remaining) return ReportViewLimit(location);

        string width = Number(chart.Width);
        string height = Number(chart.Height);
        string label = WebUtility.HtmlEncode(chart.Label);
        var html = new StringBuilder(chart.ImageUri.Length + 800);
        html.Append("<figure data-officeimo-chart><svg viewBox=\"0 0 ").Append(width).Append(' ').Append(height)
            .Append("\" width=\"").Append(width).Append("pt\" height=\"").Append(height)
            .Append("pt\" role=\"img\" aria-label=\"").Append(label).Append("\"><image width=\"").Append(width)
            .Append("\" height=\"").Append(height).Append("\" href=\"").Append(chart.ImageUri)
            .Append("\"></image><rect data-officeimo-chart-highlight visibility=\"hidden\"></rect></svg>")
            .Append("<div data-officeimo-chart-tools role=\"group\" aria-label=\"Chart navigation\" hidden>")
            .Append("<button type=\"button\" data-chart-action=\"in\">Zoom in</button>")
            .Append("<button type=\"button\" data-chart-action=\"out\">Zoom out</button>")
            .Append("<button type=\"button\" data-chart-action=\"reset\">Reset view</button>")
            .Append("<button type=\"button\" data-chart-action=\"previous\" hidden>Previous data point</button>")
            .Append("<button type=\"button\" data-chart-action=\"next\" hidden>Next data point</button>")
            .Append("</div><output data-officeimo-chart-value aria-live=\"polite\" hidden></output>");
        int markCount = AppendMarks(html, chart, remaining - 128, out bool limited);
        html.Append("</figure>\n");
        if (html.Length > remaining || (markCount == 0 && limited)) return ReportViewLimit(location);
        if (limited) Report("notebook-chart-inspection-limit", location, "Some saved chart values exceed the HTML inspection budget.");
        _views++;
        _characters += html.Length;
        return new NotebookGeneratedHtmlBlock(html.ToString());
    }

    private int AppendMarks(StringBuilder output, NotebookChartPreview chart, int maximumLength, out bool limited) {
        limited = false;
        if (chart.Marks is not { Count: > 0 } marks) return 0;
        int count = 0;
        output.Append("<template data-officeimo-chart-marks>[");
        foreach (NotebookChartMark mark in marks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (mark.Width <= 0 || mark.Height <= 0 || mark.Label.Length == 0 ||
                !Finite(mark.X) || !Finite(mark.Y) || !Finite(mark.Width) || !Finite(mark.Height)) continue;
            double x = Math.Max(0, mark.X);
            double y = Math.Max(0, mark.Y);
            double width = Math.Min(chart.Width, mark.X + mark.Width) - x;
            double height = Math.Min(chart.Height, mark.Y + mark.Height) - y;
            if (width < 0.001 || height < 0.001) continue;
            string label = mark.Label.Length <= MaxLabelCharacters ? mark.Label : mark.Label.Substring(0, MaxLabelCharacters - 1) + "…";
            string encoded = JsonEncodedText.Encode(label).ToString();
            if (count >= MaxMarksPerView || output.Length + encoded.Length + 96 > maximumLength) {
                limited = true;
                break;
            }
            if (count++ > 0) output.Append(',');
            output.Append('[').Append(Number(x)).Append(',').Append(Number(y)).Append(',')
                .Append(Number(width)).Append(',').Append(Number(height)).Append(",\"").Append(encoded).Append("\"]");
        }
        output.Append("]</template>");
        return count;
    }

    internal string AddRuntime(string html, bool fragment) {
        if (_views == 0) return html;
        if (fragment) return "<div data-officeimo-chart-export>" + html + Assets.Value + "</div>";
        int bodyEnd = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
        return bodyEnd >= 0 ? html.Insert(bodyEnd, Assets.Value) : html + Assets.Value;
    }

    private IMarkdownBlock? ReportViewLimit(string location) {
        Report("notebook-chart-interaction-limit", location, "The chart exceeds the HTML interaction budget. A static representation was used when available.");
        return null;
    }

    private void Report(string code, string location, string message) =>
        _diagnostics.Add(new OfficeDiagnostic(code, OfficeDiagnosticSeverity.Warning, message, location,
            "Jupyter", OfficeConversionLossKind.Approximation));

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static string Number(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

    private static string ReadAsset(string extension) {
        using Stream stream = typeof(NotebookHtmlCharts).Assembly.GetManifestResourceStream("OfficeIMO.Notebook.Html.NotebookHtmlCharts." + extension)
            ?? throw new InvalidOperationException("The OfficeIMO notebook chart resource is unavailable.");
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
