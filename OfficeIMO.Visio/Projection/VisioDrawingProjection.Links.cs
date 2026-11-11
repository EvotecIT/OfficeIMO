using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal sealed partial class VisioDrawingProjection {
    /// <summary>Named destination a document-wide consumer places at the top of each projected page.</summary>
    internal static string PageDestinationName(int ordinal) =>
        "visio-page-" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture);

    private IReadOnlyDictionary<string, int>? _pageOrdinals;

    /// <summary>Enables links to other pages when every page of the document is projected in one operation.</summary>
    internal void UseDocumentPages(VisioDocument document) {
        var ordinals = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int index = 0; index < document.Pages.Count; index++) {
            VisioPage page = document.Pages[index];
            foreach (string? name in new[] { page.Name, page.NameU })
                if (!string.IsNullOrEmpty(name) && !ordinals.ContainsKey(name!)) ordinals.Add(name!, index + 1);
        }
        _pageOrdinals = ordinals;
    }

    private void ProjectLinks(OfficeDrawing drawing, IList<VisioHyperlink> hyperlinks, IReadOnlyList<OfficePoint> area, string location) {
        if (area.Count == 0) return;
        ProjectLinks(drawing, hyperlinks, new[] { (area.Min(point => point.X), area.Min(point => point.Y), area.Max(point => point.X), area.Max(point => point.Y)) }, location);
    }

    /// <summary>Covers a route with short padded pieces so a diagonal link does not claim its whole bounding box.</summary>
    private static List<(double Left, double Top, double Right, double Bottom)> RouteLinkAreas(IReadOnlyList<OfficePoint> route, double reach) {
        const int MaximumAreas = 256;
        double total = 0D;
        for (int index = 1; index < route.Count; index++) total += Distance(route[index - 1], route[index]);
        double piece = Math.Max(reach * 4D, total / MaximumAreas);
        var areas = new List<(double, double, double, double)>();
        for (int index = 1; index < route.Count; index++) {
            OfficePoint from = route[index - 1], to = route[index];
            int count = Math.Max(1, (int)Math.Ceiling(Distance(from, to) / piece));
            for (int step = 0; step < count; step++) {
                double t0 = (double)step / count, t1 = (double)(step + 1) / count;
                double x0 = from.X + (to.X - from.X) * t0, y0 = from.Y + (to.Y - from.Y) * t0;
                double x1 = from.X + (to.X - from.X) * t1, y1 = from.Y + (to.Y - from.Y) * t1;
                areas.Add((Math.Min(x0, x1) - reach, Math.Min(y0, y1) - reach, Math.Max(x0, x1) + reach, Math.Max(y0, y1) + reach));
            }
        }
        return areas;
    }

    private static double Distance(OfficePoint first, OfficePoint second) =>
        Math.Sqrt(((second.X - first.X) * (second.X - first.X)) + ((second.Y - first.Y) * (second.Y - first.Y)));

    private void ProjectLinks(OfficeDrawing drawing, IList<VisioHyperlink> hyperlinks,
        IReadOnlyList<(double Left, double Top, double Right, double Bottom)> areas, string location) {
        if (hyperlinks.Count == 0 || areas.Count == 0) return;
        // One area activates one target: use the default row, else the first row with a target.
        VisioHyperlink? link = hyperlinks.FirstOrDefault(item => item.Default && !item.Invisible) ?? hyperlinks.FirstOrDefault(item => !item.Invisible);
        string? uri = link == null ? null : ResolveLinkTarget(link);
        if (uri != null && OfficeDrawingLinkPolicy.TryNormalize(uri, out string safe)) {
            string? description = string.IsNullOrWhiteSpace(link!.Description) ? null : link.Description;
            foreach ((double Left, double Top, double Right, double Bottom) area in areas) {
                double left = Math.Max(0D, area.Left), top = Math.Max(0D, area.Top);
                double right = Math.Min(drawing.Width, area.Right), bottom = Math.Min(drawing.Height, area.Bottom);
                if (right - left > 0D && bottom - top > 0D) drawing.AddLink(safe, left, top, right - left, bottom - top, description);
            }
            if (hyperlinks.Count(item => !item.Invisible) > 1)
                _report.Add("VISIO_DRAWING_LINKS", "Only the default hyperlink of a shape with several hyperlink rows is clickable.", OfficeConversionLossKind.Omission, location);
            return;
        }
        _report.Add("VISIO_DRAWING_LINKS", "The hyperlink target is not a web, mail, telephone or projected page target and is not clickable.", OfficeConversionLossKind.Omission, location);
    }

    private string? ResolveLinkTarget(VisioHyperlink link) {
        string? address = link.Address?.Trim();
        string? subAddress = link.SubAddress?.Trim();
        if (!string.IsNullOrEmpty(address)) {
            bool web = address!.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || address.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
            return web && !string.IsNullOrEmpty(subAddress) && address.IndexOf('#') < 0 ? address + "#" + Uri.EscapeDataString(subAddress!) : address;
        }
        if (string.IsNullOrEmpty(subAddress) || _pageOrdinals == null) return null;
        // A sub-address names a page, optionally followed by /shape.
        string pageName = subAddress!.Split('/')[0];
        return _pageOrdinals.TryGetValue(pageName, out int ordinal) ? "#" + PageDestinationName(ordinal) : null;
    }
}
