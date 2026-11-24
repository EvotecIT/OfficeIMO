using System.Text;
using System.Text.Json;

namespace OfficeIMO.Notebook.Html;

internal sealed partial class NotebookHtmlWidgets {
    private const int MaxLinks = 512;
    private static readonly Lazy<string> Runtime = new(ReadRuntime);

    internal string AddRuntime(string html, bool fragment) {
        if (_values.Count == 0) return html;
        var links = new StringBuilder();
        int linkCount = 0;
        foreach (string id in _models.Keys) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!TryReadModel(id, out string name, out JsonElement state)
                || name is not ("LinkModel" or "DirectionalLinkModel")) continue;
            bool bothWays = name == "LinkModel";
            bool sourceValid = TryReadEndpoint(Property(state, "source"), out string source, out string sourceTrait);
            bool targetValid = TryReadEndpoint(Property(state, "target"), out string target, out string targetTrait);
            bool sourceRendered = _values.TryGetValue(source, out WidgetValue from);
            bool targetRendered = _values.TryGetValue(target, out WidgetValue to);
            if (!sourceRendered && !targetRendered) continue;
            if (!sourceValid || !targetValid || !sourceRendered || !targetRendered
                || from.Trait != sourceTrait || to.Trait != targetTrait || !CanLink(from, to)
                || bothWays && !CanLink(to, from) || linkCount >= MaxLinks) {
                Report("notebook-widget-link-unsupported", "metadata.widgets",
                    "A saved widget link does not have two visible controls with compatible supported values.");
                continue;
            }
            linkCount++;
            links.Append("<span hidden data-officeimo-widget-link-from=\"").Append(Encode(source))
                .Append("\" data-officeimo-widget-link-to=\"").Append(Encode(target)).Append('"');
            if (bothWays) links.Append(" data-officeimo-widget-link-reverse");
            links.Append("></span>");
        }

        if (linkCount == 0 && !_needsRuntime) return html;
        links.Append("<script>\n").Append(Runtime.Value).Append("\n</script>");
        if (fragment) return "<div class=\"officeimo-widget-export\">" + html + links.ToString() + "</div>";
        int bodyEnd = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
        return bodyEnd >= 0 ? html.Insert(bodyEnd, links.ToString()) : html + links.ToString();
    }

    private static bool TryReadEndpoint(JsonElement endpoint, out string id, out string trait) {
        id = string.Empty;
        trait = string.Empty;
        if (endpoint.ValueKind != JsonValueKind.Array || endpoint.GetArrayLength() != 2) return false;
        return TryReference(endpoint[0], out id) && TryText(endpoint[1], 32, out trait);
    }

    private static bool CanLink(WidgetValue from, WidgetValue to) =>
        from.Kind == to.Kind && from.Min >= to.Min && from.Max <= to.Max
        && (!to.Integer || from.Integer)
        && (to.Step == 0 || IsGridPoint(from.Min, to.Min, to.Step)
            && (from.Min == from.Max || from.Step > 0 && IsGridPoint(from.Step, 0, to.Step)));

    private static bool IsGridPoint(double value, double origin, double spacing) {
        double index = (value - origin) / spacing;
        return Math.Abs(index) <= 1e12 && Math.Abs(index - Math.Round(index)) <= 1e-9;
    }

    private static string ReadRuntime() {
        using Stream stream = typeof(NotebookHtmlWidgets).Assembly.GetManifestResourceStream("OfficeIMO.Notebook.Html.NotebookHtmlWidgets.js")
            ?? throw new InvalidOperationException("The OfficeIMO notebook controls resource is unavailable.");
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }
}
