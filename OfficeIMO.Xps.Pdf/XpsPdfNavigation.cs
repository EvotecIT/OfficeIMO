using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Xps;

internal sealed class XpsPdfNavigation {
    private sealed class Outline {
        internal string Title = "";
        internal int Level, Order;
        internal string? Name, Uri;
    }
    private readonly Dictionary<int, List<Outline>> _outlines = new();
    private readonly int _pageCount;

    internal XpsPdfNavigation(XpsDocument document, CancellationToken token) {
        _pageCount = document.Pages.Count;
        int order = 0;
        foreach (var structure in document.ReadDocumentStructures(token, activeOnly: true)) {
            foreach (var entry in structure.Markup.Elements(document.StructureNamespace + "DocumentStructure.Outline")
                .Elements(document.StructureNamespace + "DocumentOutline").Elements(document.StructureNamespace + "OutlineEntry")) {
                token.ThrowIfCancellationRequested();
                if (++order > 100000) throw new InvalidDataException("XPS outline entry limit exceeded.");
                string title = (string?)entry.Attribute("Description") ?? "";
                if (string.IsNullOrWhiteSpace(title)) throw new NotSupportedException("PDF outline entries require a nonempty description.");
                if (!int.TryParse((string?)entry.Attribute("OutlineLevel") ?? "1", NumberStyles.None, CultureInfo.InvariantCulture, out int level)
                    || level < 1 || level > PdfPageCanvas.MaximumOutlineLevel) throw new NotSupportedException("XPS outline depth exceeds the PDF outline profile.");
                var reference = document.ResolveNavigation(structure.Part, (string?)entry.Attribute("OutlineTarget") ?? "");
                if (!reference.HasValue) throw new NotSupportedException("Unresolved or unsafe XPS outline target.");
                int page = reference.Value.Uri == null ? reference.Value.PageIndex : 0;
                if (!_outlines.TryGetValue(page, out var entries)) _outlines.Add(page, entries = new List<Outline>());
                entries.Add(new Outline { Title = title, Level = level, Order = order, Name = reference.Value.Name, Uri = reference.Value.Uri });
            }
        }
    }

    internal string MapLinks(string svg, int pageIndex, CancellationToken token) {
        var xml = XElement.Parse(svg);
        foreach (var link in xml.Descendants(XName.Get("a", "http://www.w3.org/2000/svg"))) {
            token.ThrowIfCancellationRequested();
            var href = link.Attribute("href");
            if (href == null) continue;
            string value = href.Value;
            if (Uri.TryCreate(value, UriKind.Absolute, out _)) continue;
            int page = pageIndex; string? name = null;
            int hash = value.IndexOf('#');
            string path = hash < 0 ? value : value.Substring(0, hash);
            if (hash >= 0) {
                string fragment = value.Substring(hash + 1);
                if (fragment.Length != 0 && !fragment.StartsWith("xps-", StringComparison.Ordinal)) throw new InvalidDataException("Invalid projected XPS link.");
                name = fragment.Length == 0 ? null : fragment.Substring(4);
            }
            if (path.Length > 0) {
                if (!path.StartsWith("page-", StringComparison.Ordinal) || !path.EndsWith(".svg", StringComparison.Ordinal)
                    || !int.TryParse(path.Substring(5, path.Length - 9), NumberStyles.None, CultureInfo.InvariantCulture, out int number)
                    || number < 1 || number > _pageCount) throw new InvalidDataException("Invalid projected XPS page link.");
                page = number - 1;
            }
            href.Value = "#" + Destination(page, name);
        }
        return xml.ToString(SaveOptions.DisableFormatting);
    }

    internal void AddDestinations(PdfPageCanvas canvas, XpsSvgResult projection, int pageIndex, double width, double height, CancellationToken token) {
        canvas.NamedDestination(Destination(pageIndex, null), 0, 0);
        var locations = new Dictionary<string, OfficePoint>(StringComparer.Ordinal);
        foreach (var target in projection.Targets) {
            token.ThrowIfCancellationRequested();
            if (locations.ContainsKey(target.Name)) continue;
            var point = target.TopLeft;
            point = new OfficePoint(Math.Max(0, Math.Min(width, point.X * .75)), Math.Max(0, Math.Min(height, point.Y * .75)));
            locations.Add(target.Name, point);
            canvas.NamedDestination(Destination(pageIndex, target.Name), point.X, point.Y);
        }
        if (!_outlines.TryGetValue(pageIndex, out var outlines)) return;
        foreach (var outline in outlines) {
            token.ThrowIfCancellationRequested();
            var point = outline.Name != null && locations.TryGetValue(outline.Name, out var named) ? named : new OfficePoint(0, 0);
            canvas.OutlineNavigation(outline.Title, outline.Level, point.X, point.Y, outline.Uri, outline.Order);
        }
    }

    private static string Destination(int pageIndex, string? name) => "xps-page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture)
        + (name == null ? "" : "-name-" + Convert.ToBase64String(Encoding.UTF8.GetBytes(name)));
}
