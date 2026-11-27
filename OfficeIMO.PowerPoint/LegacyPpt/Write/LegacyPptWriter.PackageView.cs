using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    /// <summary>
    /// Reading helpers for the package view the writer captures: attribute checks that match what the binary format can
    /// hold, the parts that own a slide's theme and background, and the way back from a node of the view to the same
    /// element of the live document.
    /// </summary>
    internal static partial class LegacyPptWriter {
        private const int Pml = LegacyPptXml.P;
        private const int Dml = LegacyPptXml.A;

        /// <summary>The view part that holds a slide.</summary>
        private static LegacyPptPartView SlidePartView(LegacyPptPackageView view, PowerPointSlide slide) =>
            view.Part(slide.PartRef.Name);

        /// <summary>The root element's common slide data (<c>p:cSld</c>), or none.</summary>
        private static OfficeXmlNode CommonSlideData(LegacyPptPartView part) => part.Root.Child(Pml, "cSld");

        /// <summary>The part's shape tree (<c>p:cSld/p:spTree</c>), or none.</summary>
        private static OfficeXmlNode ShapeTree(LegacyPptPartView part) => CommonSlideData(part).Child(Pml, "spTree");

        /// <summary>The theme part's <c>a:themeElements</c> child named <paramref name="localName"/>, or none.</summary>
        private static OfficeXmlNode ThemeElement(LegacyPptPartView themePart, string localName) =>
            themePart.Root.Path(Dml, "themeElements", localName);

        /// <summary>A theme override part's root child named <paramref name="localName"/>, or none.</summary>
        private static OfficeXmlNode ThemeOverrideElement(LegacyPptPartView themeOverridePart, string localName) =>
            themeOverridePart.Root.Child(Dml, localName);

        /// <summary>
        /// Whether the element carries an attribute the SDK reported: namespace declarations and markup-compatibility
        /// attributes are not attributes there.
        /// </summary>
        private static bool HasAttributes(OfficeXmlNode element) => AttributeLocalNames(element).Count > 0;

        /// <summary>The local names of the element's attributes, without namespace declarations or compatibility markup.</summary>
        private static List<string> AttributeLocalNames(OfficeXmlNode element) {
            var names = new List<string>();
            if (element.Tree is not { } tree) return names;
            int count = tree.AttributeCountOf(element.Node);
            for (int index = 0; index < count; index++) {
                tree.GetAttributeAt(element.Node, index, out int namespaceId, out string localName, out _);
                if (namespaceId == OfficeXmlNamespaces.Xmlns
                    || namespaceId == OfficeXmlNamespaces.None && string.Equals(localName, "xmlns", StringComparison.Ordinal)
                    || namespaceId == OfficeXmlNamespaces.MarkupCompatibility) continue;
                names.Add(localName);
            }
            return names;
        }

        /// <summary>Whether the element has child elements.</summary>
        private static bool HasChildElements(OfficeXmlNode element) => !element.FirstChild.IsNone;

        /// <summary>
        /// The element of the live document at the same place as <paramref name="node"/> of the view part
        /// <paramref name="partName"/>: the inverse of <see cref="LegacyPptPackageView.Resolve"/>, for callees that still
        /// read the live document. A node already read from the live document is returned as it is.
        /// </summary>
        private static OfficeXmlNode ToLiveNode(PowerPointDocument presentation, string partName, OfficeXmlNode node) {
            if (node.IsNone || node.IsSdk) return node;
            var path = new Stack<int>();
            OfficeXmlNode current = node;
            while (true) {
                OfficeXmlNode parent = current.Parent;
                if (parent.IsNone) break;
                int index = 0;
                for (OfficeXmlNode sibling = current.PreviousSibling; !sibling.IsNone; sibling = sibling.PreviousSibling) index++;
                path.Push(index);
                current = parent;
            }
            OfficeXmlNode live = PowerPointPartRef.Package(presentation, partName).Root;
            if (live.IsNone || !string.Equals(live.LocalName, current.LocalName, StringComparison.Ordinal)) return default;
            while (path.Count > 0) {
                int index = path.Pop();
                OfficeXmlNode child = live.FirstChild;
                for (int position = 0; position < index && !child.IsNone; position++) child = child.NextSibling;
                if (child.IsNone) return default;
                live = child;
            }
            return string.Equals(live.LocalName, node.LocalName, StringComparison.Ordinal) ? live : default;
        }

        /// <summary>The placeholder (<c>p:nvXxPr/p:nvPr/p:ph</c>) of a shape, picture or graphic frame element, or none.</summary>
        private static OfficeXmlNode PlaceholderOf(OfficeXmlNode element) {
            string? container = element.NamespaceId != Pml ? null : element.LocalName switch {
                "sp" => "nvSpPr",
                "pic" => "nvPicPr",
                "graphicFrame" => "nvGraphicFramePr",
                _ => null
            };
            return container == null ? default : element.Path(Pml, container, "nvPr", "ph");
        }

        /// <summary>The binary placeholder kind of a placeholder on a slide, as the slide atom lists it.</summary>
        private static byte MapSlidePlaceholder(string? type, bool vertical) => type switch {
            null => 0,
            "title" => vertical ? (byte)0x11 : (byte)0x0D,
            "ctrTitle" => 0x0F,
            "subTitle" => 0x10,
            "body" => vertical ? (byte)0x12 : (byte)0x0E,
            "obj" => vertical ? (byte)0x19 : (byte)0x13,
            "chart" => 0x14,
            "tbl" => 0x15,
            "clipArt" => 0x16,
            "dgm" => 0x17,
            "media" => 0x18,
            "pic" => 0x1A,
            "sldImg" => 0x0B,
            "dt" => 0x07,
            "sldNum" => 0x08,
            "ftr" => 0x09,
            "hdr" => 0x0A,
            _ => 0
        };
    }
}
