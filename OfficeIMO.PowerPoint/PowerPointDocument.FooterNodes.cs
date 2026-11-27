using OfficeIMO.OpenXml.Internal;
using Ns = OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>Whether a run (<c>a:r</c>) is the marked slide-count run that follows a footer's slide number.</summary>
        internal static bool IsFooterTotal(OfficeXmlNode run) {
            foreach (OfficeXmlNode extension in run.Child(Ns.DrawingML, "rPr").Child(Ns.DrawingML, "extLst").Children(Ns.DrawingML, "ext")) {
                if (extension.Attribute("uri") == FooterTotalUri) return true;
            }
            return false;
        }

        /// <summary>Whether an inline element of a paragraph is a slide-number field (<c>a:fld type="slidenum"</c>).</summary>
        internal static bool IsSlideNumberField(OfficeXmlNode inline) =>
            inline.Is(Ns.DrawingML, "fld") && inline.Attribute("type") == "slidenum";
    }
}
