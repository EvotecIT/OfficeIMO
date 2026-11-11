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
        if (hyperlinks.Count == 0 || area.Count == 0) return;
        double left = Math.Max(0D, area.Min(point => point.X)), top = Math.Max(0D, area.Min(point => point.Y));
        double right = Math.Min(drawing.Width, area.Max(point => point.X)), bottom = Math.Min(drawing.Height, area.Max(point => point.Y));
        // One rectangle activates one target: use the default row, else the first row with a target.
        VisioHyperlink? link = hyperlinks.FirstOrDefault(item => item.Default && !item.Invisible) ?? hyperlinks.FirstOrDefault(item => !item.Invisible);
        string? uri = link == null ? null : ResolveLinkTarget(link);
        if (uri != null && right - left > 0D && bottom - top > 0D && OfficeDrawingLinkPolicy.TryNormalize(uri, out string safe)) {
            drawing.AddLink(safe, left, top, right - left, bottom - top, string.IsNullOrWhiteSpace(link!.Description) ? null : link.Description);
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
