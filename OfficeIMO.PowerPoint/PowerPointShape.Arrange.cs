using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>
        ///     Moves the shape to the front (top) of the z-order within its parent.
        /// </summary>
        public void BringToFront() {
            OpenXmlElement? parent = Element.Parent;
            if (parent == null) {
                return;
            }

            Element.Remove();
            parent.Append(Element);
        }

        /// <summary>
        ///     Moves the shape to the back (bottom) of the z-order within its parent.
        /// </summary>
        public void SendToBack() {
            OpenXmlElement? parent = Element.Parent;
            if (parent == null) {
                return;
            }

            Element.Remove();

            OpenXmlElement? insertBefore = null;
            foreach (OpenXmlElement child in parent.ChildElements) {
                if (child is NonVisualGroupShapeProperties || child is GroupShapeProperties) {
                    continue;
                }
                insertBefore = child;
                break;
            }

            if (insertBefore != null) {
                parent.InsertBefore(Element, insertBefore);
            } else {
                parent.Append(Element);
            }
        }
    }
}
