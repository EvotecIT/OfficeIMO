using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>
        ///     Moves the shape to the front (top) of the z-order within its parent.
        /// </summary>
        public void BringToFront() {
            OfficeXmlNode element = XmlElement;
            OfficeXmlNode parent = element.Parent;
            if (parent.IsNone) {
                return;
            }

            element.Remove();
            parent.Append(element);
            OwnerPartRef.MarkEdited();
        }

        /// <summary>
        ///     Moves the shape to the back (bottom) of the z-order within its parent.
        /// </summary>
        public void SendToBack() {
            OfficeXmlNode element = XmlElement;
            OfficeXmlNode parent = element.Parent;
            if (parent.IsNone) {
                return;
            }

            element.Remove();

            // The group's own properties stay first.
            OfficeXmlNode insertBefore = default;
            foreach (OfficeXmlNode child in parent.Children()) {
                if (child.Is(OfficeXmlNamespaces.PresentationML, "nvGrpSpPr") || child.Is(OfficeXmlNamespaces.PresentationML, "grpSpPr")) {
                    continue;
                }
                insertBefore = child;
                break;
            }

            parent.InsertBefore(element, insertBefore);
            OwnerPartRef.MarkEdited();
        }
    }
}
