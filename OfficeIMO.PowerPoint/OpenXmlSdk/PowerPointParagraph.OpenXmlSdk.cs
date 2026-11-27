using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The SDK side of a paragraph: a constructor over an SDK paragraph and SDK views of it, for code that still works on
    /// SDK elements. Reaching an SDK view of a paragraph held in a package tree opens the SDK.
    /// </summary>
    public partial class PowerPointParagraph {
        internal PowerPointParagraph(A.Paragraph paragraph, SlidePart? slidePart = null, OpenXmlPartContainer? ownerPart = null)
            : this(OfficeXmlNode.FromSdk(paragraph), PowerPointPartRef.FromSdk(slidePart ?? ownerPart as SlidePart), PowerPointPartRef.FromSdk(ownerPart as OpenXmlPart ?? slidePart)) {
        }

        /// <summary>The SDK paragraph.</summary>
        internal A.Paragraph Paragraph {
            get {
                OfficeXmlNode element = Element;
                if (element.SdkElement is A.Paragraph paragraph) return paragraph;
                if (_packageSlide is { RemovedFromPackage: true } || _owningSlide is { RemovedFromPackage: true }) throw new InvalidOperationException("The slide is no longer part of the presentation.");
                if (!_standalone && element.Tree != null && !_packagePart.IsNone) return (A.Paragraph)PowerPointSlide.BindPackageNode(_packagePart, element.Tree, element.Node);
                throw new InvalidOperationException("The paragraph is not held in an SDK part.");
            }
        }

        /// <summary>The SDK part that holds the paragraph.</summary>
        internal OpenXmlPartContainer? OwnerPart {
            get {
                PowerPointPartRef owner = OwnerPartRef;
                if (owner.IsNone) return null;
                if (owner.Current.SdkPart == null && owner.Document is { } document) _ = document.OpenXmlDocument;
                return owner.Current.SdkPart as OpenXmlPart;
            }
        }
    }
}
