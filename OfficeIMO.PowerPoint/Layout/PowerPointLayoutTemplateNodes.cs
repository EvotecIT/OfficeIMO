using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;
using Ns = OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Reads a template's masters, layouts and placeholder geometry for the layout system, from the parts wherever they
    /// are held.
    /// </summary>
    internal static class PowerPointLayoutTemplateNodes {
        /// <summary>The layout part at <paramref name="layoutIndex"/> of the master at <paramref name="masterIndex"/>.</summary>
        internal static PowerPointPartRef LayoutPart(PowerPointDocument document, int masterIndex, int layoutIndex) {
            PowerPointPartRef[] masters = document.LayoutPresentationPartRef.RelatedParts(PowerPointPartRef.SlideMasterRelationshipType).ToArray();
            ArgumentOutOfRangeException.ThrowIfNegative(masterIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(masterIndex, masters.Length);
            PowerPointPartRef[] layouts = masters[masterIndex].RelatedParts(PowerPointPartRef.SlideLayoutRelationshipType).ToArray();
            ArgumentOutOfRangeException.ThrowIfNegative(layoutIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(layoutIndex, layouts.Length);
            return layouts[layoutIndex];
        }

        /// <summary>The shape tree (<c>p:cSld/p:spTree</c>) of a slide, layout or master root.</summary>
        internal static OfficeXmlNode ShapeTree(OfficeXmlNode root) =>
            root.Child(Ns.PresentationML, "cSld").Child(Ns.PresentationML, "spTree");

        /// <summary>The placeholder (<c>p:ph</c>) of a shape, picture or graphic frame, or none.</summary>
        internal static OfficeXmlNode Placeholder(OfficeXmlNode element) {
            string? container = element.NamespaceId != Ns.PresentationML ? null : element.LocalName switch {
                "sp" => "nvSpPr",
                "pic" => "nvPicPr",
                "graphicFrame" => "nvGraphicFramePr",
                _ => null
            };
            return container == null ? default
                : element.Child(Ns.PresentationML, container).Child(Ns.PresentationML, "nvPr").Child(Ns.PresentationML, "ph");
        }

        /// <summary>The name (<c>p:cNvPr/@name</c>) of a shape, or empty.</summary>
        internal static string ShapeName(OfficeXmlNode shape) =>
            shape.Child(Ns.PresentationML, "nvSpPr").Child(Ns.PresentationML, "cNvPr").Attribute("name") ?? string.Empty;

        /// <summary>The placeholder type token, <c>body</c> when it is omitted.</summary>
        internal static string PlaceholderType(OfficeXmlNode placeholder) => placeholder.Attribute("type") ?? "body";

        /// <summary>The placeholder index, 0 when it is omitted.</summary>
        internal static uint PlaceholderIndex(OfficeXmlNode placeholder) => placeholder.UInt32Attribute("idx") ?? 0U;

        /// <summary>The bounds a shape, picture or graphic frame states itself, or null when its transform is incomplete.</summary>
        internal static PowerPointLayoutBox? ElementBounds(OfficeXmlNode element) {
            if (element.NamespaceId != Ns.PresentationML) return null;
            OfficeXmlNode transform = element.LocalName switch {
                "sp" or "pic" => element.Child(Ns.PresentationML, "spPr").Child(Ns.DrawingML, "xfrm"),
                "graphicFrame" => element.Child(Ns.PresentationML, "xfrm"),
                _ => default
            };
            return TransformBounds(transform);
        }

        /// <summary>The bounds of an <c>a:xfrm</c> or <c>p:xfrm</c>, or null when its offset or extents are incomplete.</summary>
        internal static PowerPointLayoutBox? TransformBounds(OfficeXmlNode transform) {
            OfficeXmlNode offset = transform.Child(Ns.DrawingML, "off"), extents = transform.Child(Ns.DrawingML, "ext");
            long? x = offset.Int64Attribute("x"), y = offset.Int64Attribute("y");
            long? cx = extents.Int64Attribute("cx"), cy = extents.Int64Attribute("cy");
            if (x == null || y == null || cx == null || cy == null) return null;
            return new PowerPointLayoutBox(OfficeLength.Emus(x.Value), OfficeLength.Emus(y.Value), OfficeLength.Emus(cx.Value), OfficeLength.Emus(cy.Value));
        }

        /// <summary>
        /// Resolves an omitted placeholder transform through the layout and then the master, as slides inherit it: the
        /// layout placeholder with the same index, then the master placeholder of the same type.
        /// </summary>
        internal static PowerPointLayoutBox? ExportPlaceholderBounds(PowerPointPartRef layoutPart, OfficeXmlNode placeholder) {
            string type = PlaceholderType(placeholder);
            uint index = PlaceholderIndex(placeholder);
            OfficeXmlNode layoutShape = default;
            foreach (OfficeXmlNode element in ShapeTree(layoutPart.Root).Children()) {
                OfficeXmlNode candidate = Placeholder(element);
                if (candidate.Exists && PlaceholderIndex(candidate) == index) {
                    layoutShape = element;
                    break;
                }
            }
            if (layoutShape.Exists) {
                PowerPointLayoutBox? local = ElementBounds(layoutShape);
                if (local.HasValue) return local;
                type = PlaceholderType(Placeholder(layoutShape));
            }
            // Layout placeholders match the master's semantic type, rather than the slide/layout index (which commonly
            // differs on title/content layouts).
            PowerPointPartRef master = layoutPart.IsNone ? default : layoutPart.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType);
            foreach (OfficeXmlNode element in ShapeTree(master.Root).Children()) {
                OfficeXmlNode candidate = Placeholder(element);
                if (candidate.Exists && SamePlaceholderType(PlaceholderType(candidate), type)) return ElementBounds(element);
            }
            return null;
        }

        private static bool SamePlaceholderType(string left, string right) => left == right ||
            IsTitle(left) && IsTitle(right);

        internal static bool IsTitle(string type) => type is "title" or "ctrTitle";
    }
}
