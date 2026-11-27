using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    private const int PresentationNamespace = OfficeXmlNamespaces.PresentationML;

    /// <summary>
    /// Resolves an omitted placeholder transform (<paramref name="placeholder"/> is its <c>p:ph</c>) through the layout,
    /// matched by index, and then the master, matched by type.
    /// </summary>
    internal static PowerPointLayoutBox? GetExportPlaceholderBounds(PowerPointPartRef layout, OfficeXmlNode placeholder) {
        string type = placeholder.Attribute("type") ?? "body";
        uint index = placeholder.UInt32Attribute("idx") ?? 0U;
        foreach (OfficeXmlNode element in ShapeTreeOf(layout).Children()) {
            OfficeXmlNode candidate = PlaceholderOf(element);
            if (candidate.IsNone || (candidate.UInt32Attribute("idx") ?? 0U) != index) continue;
            PowerPointLayoutBox? local = ElementBounds(element);
            if (local.HasValue) return local;
            type = candidate.Attribute("type") ?? "body";
            break;
        }
        // Layout placeholders match the master's semantic type, rather than the
        // slide/layout index (which commonly differs on title/content layouts).
        foreach (OfficeXmlNode element in ShapeTreeOf(layout.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType)).Children()) {
            OfficeXmlNode candidate = PlaceholderOf(element);
            if (!candidate.IsNone && SamePlaceholderType(candidate.Attribute("type") ?? "body", type)) return ElementBounds(element);
        }
        return null;
    }

    private static OfficeXmlNode ShapeTreeOf(PowerPointPartRef part) =>
        part.IsNone ? default : part.Root.Child(PresentationNamespace, "cSld").Child(PresentationNamespace, "spTree");

    // The p:ph of a shape, picture or graphic frame, or none.
    private static OfficeXmlNode PlaceholderOf(OfficeXmlNode element) {
        string? container = element.NamespaceId != PresentationNamespace ? null : element.LocalName switch {
            "sp" => "nvSpPr",
            "pic" => "nvPicPr",
            "graphicFrame" => "nvGraphicFramePr",
            _ => null
        };
        return container == null ? default
            : element.Child(PresentationNamespace, container).Child(PresentationNamespace, "nvPr").Child(PresentationNamespace, "ph");
    }

    // The offset and extents of a shape's or picture's a:xfrm, or a graphic frame's p:xfrm, when all four are set.
    private static PowerPointLayoutBox? ElementBounds(OfficeXmlNode element) {
        OfficeXmlNode transform = element.LocalName == "graphicFrame"
            ? element.Child(PresentationNamespace, "xfrm")
            : element.Child(PresentationNamespace, "spPr").Child(OfficeXmlNamespaces.DrawingML, "xfrm");
        OfficeXmlNode offset = transform.Child(OfficeXmlNamespaces.DrawingML, "off"), extents = transform.Child(OfficeXmlNamespaces.DrawingML, "ext");
        long? x = offset.Int64Attribute("x"), y = offset.Int64Attribute("y"), cx = extents.Int64Attribute("cx"), cy = extents.Int64Attribute("cy");
        if (x == null || y == null || cx == null || cy == null) return null;
        return new PowerPointLayoutBox(OfficeLength.Emus(x.Value), OfficeLength.Emus(y.Value), OfficeLength.Emus(cx.Value), OfficeLength.Emus(cy.Value));
    }

    private static bool SamePlaceholderType(string left, string right) => left == right ||
        (left is "title" or "ctrTitle") && (right is "title" or "ctrTitle");
}
