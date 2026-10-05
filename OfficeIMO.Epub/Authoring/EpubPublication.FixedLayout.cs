using System.Globalization;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private const string RenditionVocabulary = "http://www.idpf.org/vocab/rendition/#";
    private const string FixedLayoutStyleId = "officeimo-fixed-layout-canvas";

    /// <summary>
    /// Atomically configures one EPUB 3 XHTML spine document's viewport, CSS page canvas and
    /// fixed-layout presentation overrides. DOM order, identifiers and semantic relationships remain
    /// unchanged. Existing content CSS may override the canvas rules; verify rendered output in target
    /// readers. This does not paginate reflowable content or certify accessible reading order.
    /// </summary>
    public void SetFixedLayoutPage(string manifestId, EpubFixedLayoutPage page, CancellationToken cancellationToken = default) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        cancellationToken.ThrowIfCancellationRequested();
        if (PackageVersion != "3.0") throw new NotSupportedException("Fixed-layout page authoring requires EPUB 3.");
        if (page.Width < 1 || page.Height < 1) throw new ArgumentOutOfRangeException(nameof(page), "Canvas dimensions must be positive CSS pixels.");
        string orientation = page.Orientation switch {
            EpubPageOrientation.Auto => "auto", EpubPageOrientation.Landscape => "landscape",
            EpubPageOrientation.Portrait => "portrait", _ => throw new ArgumentOutOfRangeException(nameof(page))
        };
        string spread = page.Spread switch {
            EpubPageSpread.Auto => "auto", EpubPageSpread.None => "none", EpubPageSpread.Landscape => "landscape",
            EpubPageSpread.Both => "both", _ => throw new ArgumentOutOfRangeException(nameof(page))
        };
        string? side = page.Side switch {
            EpubPageSide.Auto => null, EpubPageSide.Left => "page-spread-left", EpubPageSide.Right => "page-spread-right",
            EpubPageSide.Center => "rendition:page-spread-center", _ => throw new ArgumentOutOfRangeException(nameof(page))
        };
        if (EpubVocabulary.Expand(Root, "rendition:layout") != RenditionVocabulary + "layout")
            throw new InvalidDataException("The rendition prefix is mapped to a different vocabulary.");
        XElement[] positions = RequireSection("spine").Elements(Opf + "itemref").Where(item => (string?)item.Attribute("idref") == manifestId).ToArray();
        if (positions.Length != 1) throw new InvalidDataException("Fixed-layout page authoring requires exactly one spine position for the resource.");
        XDocument document = EditableXhtml(manifestId);
        XElement head = document.Root!.Element(Html + "head") ?? throw new InvalidDataException("Content has no XHTML head.");
        if (document.Root.Element(Html + "body") == null) throw new InvalidDataException("Content has no XHTML body.");
        XElement[] viewports = head.Elements(Html + "meta").Where(meta => string.Equals((string?)meta.Attribute("name"), "viewport", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (viewports.Length > 1) throw new InvalidDataException("Content has multiple viewport declarations.");
        XElement viewport = viewports.SingleOrDefault() ?? new XElement(Html + "meta", new XAttribute("name", "viewport"));
        string width = page.Width.ToString(CultureInfo.InvariantCulture), height = page.Height.ToString(CultureInfo.InvariantCulture);
        viewport.SetAttributeValue("content", "width=" + width + ", height=" + height);
        if (viewport.Parent == null) head.Add(viewport);
        XElement[] matchingIds = document.Root.DescendantsAndSelf().Where(element => (string?)element.Attribute("id") == FixedLayoutStyleId ||
            (string?)element.Attribute(XNamespace.Xml + "id") == FixedLayoutStyleId).ToArray();
        if (matchingIds.Length > 1 || matchingIds.Length == 1 && (matchingIds[0].Name != Html + "style" ||
            matchingIds[0].Parent != head || (string?)matchingIds[0].Attribute("data-officeimo-canvas") != "1"))
            throw new InvalidDataException("The fixed-layout canvas style identifier is already used by other content.");
        XElement style = matchingIds.SingleOrDefault() ?? new XElement(Html + "style", new XAttribute("id", FixedLayoutStyleId), new XAttribute("data-officeimo-canvas", "1"));
        style.Value = "html, body { width: " + width + "px; height: " + height + "px; margin: 0; padding: 0; } body { position: relative; }";
        if (style.Parent != null) style.Remove();
        head.Add(style);
        string[] properties = Tokens((string?)positions[0].Attribute("properties")).Where(token => !IsFixedLayoutOverride(token)).Concat(new[] {
            "rendition:layout-pre-paginated", "rendition:orientation-" + orientation, "rendition:spread-" + spread
        }).Concat(side == null ? Array.Empty<string>() : new[] { side }).ToArray();
        CommitContentEdits(new Dictionary<string, XDocument>(StringComparer.Ordinal) { [manifestId] = document }, cancellationToken,
            validatePublication: true, packageEdit: root => root.Element(Opf + "spine")!.Elements(Opf + "itemref")
                .Single(item => (string?)item.Attribute("idref") == manifestId).SetAttributeValue("properties", string.Join(" ", properties)));
    }

    private bool IsFixedLayoutOverride(string token) {
        if (token == "page-spread-left" || token == "page-spread-right") return true;
        string expanded = EpubVocabulary.Expand(Root, token);
        return new[] { "layout-pre-paginated", "layout-reflowable", "orientation-auto", "orientation-landscape", "orientation-portrait",
            "spread-auto", "spread-none", "spread-landscape", "spread-portrait", "spread-both", "page-spread-left", "page-spread-right", "page-spread-center" }
            .Any(property => expanded == RenditionVocabulary + property);
    }
}
