using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>Binds package-backed slide content to the SDK elements at the same places once the SDK opens.</summary>
    public partial class PowerPointSlide {
        /// <summary>
        /// The SDK element of a node of a slide's package tree. The SDK reads the bytes the tree wrote, so the element sits
        /// at the same place: the same element child at each level down from the slide root. Comments and text the SDK
        /// keeps as nodes are not elements of the tree. Binding opens the SDK when it is not open yet.
        /// </summary>
        internal static DocumentFormat.OpenXml.OpenXmlElement BindPackageElement(PowerPointSlide slide, OfficeXmlTree tree, int node) {
            var path = new List<int>();
            for (int current = node; current != tree.Root; current = tree.ParentOf(current)) {
                // A node an edit removed from the slide has no place in it.
                if (current == OfficeXmlTree.NoNode) throw new InvalidOperationException("The element is no longer part of the slide.");
                int index = 0;
                for (int sibling = tree.PreviousSiblingOf(current); sibling != OfficeXmlTree.NoNode; sibling = tree.PreviousSiblingOf(sibling)) {
                    if (tree.KindOf(sibling) == OfficeXmlTreeNodeKind.Element) index++;
                }
                path.Add(index);
            }
            string localName = tree.LocalNameOf(node);
            DocumentFormat.OpenXml.OpenXmlElement element = slide.SlidePart.Slide
                ?? throw new InvalidOperationException("The slide the SDK opened has no content.");
            for (int level = path.Count - 1; level >= 0; level--) {
                DocumentFormat.OpenXml.OpenXmlElement? child = null;
                int remaining = path[level];
                foreach (DocumentFormat.OpenXml.OpenXmlElement candidate in element.ChildElements) {
                    if (candidate is DocumentFormat.OpenXml.OpenXmlMiscNode) continue;
                    if (remaining-- == 0) {
                        child = candidate;
                        break;
                    }
                }
                element = child ?? throw new InvalidOperationException("The element is not where the slide the SDK opened has it.");
            }
            if (!string.Equals(element.LocalName, localName, StringComparison.Ordinal)) {
                throw new InvalidOperationException("The element is not where the slide the SDK opened has it.");
            }
            return element;
        }
    }
}
