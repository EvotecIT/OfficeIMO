using OfficeIMO.Html;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static string FixedLayoutRegionCss(XDocument document, EpubFixedLayoutPage page, CancellationToken token) {
        IReadOnlyList<EpubFixedLayoutRegion> regions = page.Regions ?? throw new ArgumentException("Regions cannot be null.", nameof(page));
        if (regions.Count > 1024) throw new ArgumentOutOfRangeException(nameof(page), "A page supports at most 1024 positioned regions.");
        if (regions.Count == 0) return string.Empty;
        XElement body = document.Root!.Element(Html + "body")!;
        var targets = body.Elements().Where(element => element.Attribute("id") != null)
            .ToLookup(element => (string)element.Attribute("id")!, StringComparer.Ordinal);
        var selected = new HashSet<string>(StringComparer.Ordinal);
        var css = new StringBuilder();
        foreach (EpubFixedLayoutRegion region in regions) {
            token.ThrowIfCancellationRequested();
            if (region == null) throw new ArgumentException("Regions cannot contain null entries.", nameof(page));
            RequireText(region.ElementId, nameof(page));
            if (region.ElementId.Length > 1024) throw new ArgumentOutOfRangeException(nameof(page), "Region identifiers cannot exceed 1024 characters.");
            if (!selected.Add(region.ElementId)) throw new InvalidDataException("A page cannot position the same region twice: " + region.ElementId);
            XElement[] matches = targets[region.ElementId].ToArray();
            if (matches.Length != 1) throw new InvalidDataException("A positioned region requires one top-level body element with an HTML id: " + region.ElementId);
            XElement target = matches[0];
            if (target.Name.Namespace != Html ||
                new[] { "script", "style", "link", "meta", "template", "title", "base" }.Contains(target.Name.LocalName, StringComparer.Ordinal))
                throw new InvalidDataException("A positioned region must target a top-level XHTML body element with an HTML id: " + region.ElementId);
            if (region.Left < 0 || region.Top < 0 || region.Width <= 0 || region.Height <= 0 ||
                region.Left > page.Width - region.Width || region.Top > page.Height - region.Height)
                throw new InvalidDataException("Region '" + region.ElementId + "' must have positive dimensions and fit inside the " +
                    page.Width.ToString(CultureInfo.InvariantCulture) + " × " + page.Height.ToString(CultureInfo.InvariantCulture) + " CSS-pixel canvas.");
            css.Append("\nbody > [id=").Append(HtmlCssStringEncoder.Quote(region.ElementId)).Append("] { position: absolute; box-sizing: border-box; margin: 0; left: ")
                .Append(Pixels(region.Left)).Append("; top: ").Append(Pixels(region.Top)).Append("; width: ")
                .Append(Pixels(region.Width)).Append("; height: ").Append(Pixels(region.Height)).Append("; }");
        }
        return css.ToString();
    }

    private static string Pixels(decimal value) => value.ToString("0.############################", CultureInfo.InvariantCulture) + "px";
}
