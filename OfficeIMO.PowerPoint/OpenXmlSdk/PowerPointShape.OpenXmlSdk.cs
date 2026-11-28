using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// The SDK side of a shape: its constructor over an SDK element, binding to the SDK element, and the SDK element
    /// helpers that shape members not yet on the element handle use.
    /// </summary>
    public abstract partial class PowerPointShape {
        private OpenXmlElement? _element;
        private OpenXmlPart? _relationshipOwnerPart;

        internal PowerPointShape(OpenXmlElement element, OpenXmlPartContainer? ownerPart = null) {
            _element = element;
            RelationshipOwnerPart = ownerPart as OpenXmlPart;
        }

        // A shape of a slide whose SDK part is not open: it reads and edits its element in the slide's package tree, and
        // anything it does not do there opens the SDK and binds the shape to the same element there.
        internal PowerPointShape(PowerPointSlide slide, OfficeXmlTree tree, int node) {
            OwnerSlide = slide;
            PackageTree = tree;
            PackageNode = node;
        }

        // A shape of a layout's or master's package tree; it binds to the same element when the SDK opens.
        internal PowerPointShape(OpenXml.PowerPointPartRef part, OfficeXmlTree tree, int node) {
            PackagePart = part;
            PackageTree = tree;
            PackageNode = node;
            part.Document?.RegisterPackageBinding(this);
        }

        internal OpenXmlElement Element => _element ?? BindSdkElement();

        /// <summary>Binds a package-backed shape to its SDK element now, while the slide the SDK opened is unchanged.</summary>
        internal void BindSdk() {
            if (_element != null || PackageTree == null) return;
            try {
                BindSdkElement();
            } catch (InvalidOperationException) {
                // A shape the opened slide doesn't have where expected reports it when it is used, not while the SDK opens.
            }
        }
        internal PowerPointSlide? OwnerSlide { get; private set; }

        internal OpenXmlPart? RelationshipOwnerPart {
            get => _relationshipOwnerPart ?? (PackageTree == null ? null : !PackagePart.IsNone ? PackagePartForSdk() : OwnerSlide?.SlidePart);
            private set => _relationshipOwnerPart = value;
        }

        /// <summary>The slide tree that holds the shape while it works without the SDK, and its element there.</summary>
        internal OfficeXmlTree? PackageTree { get; }
        internal int PackageNode { get; }

        /// <summary>Whether the shape works on its slide's package tree rather than an SDK element.</summary>
        internal bool IsPackageBacked => _element == null;

        private OpenXmlElement BindSdkElement() {
            OpenXmlElement element = !PackagePart.IsNone
                ? (OpenXmlElement)PowerPointSlide.BindPackageNode(PackagePart, PackageTree!, PackageNode)
                : PowerPointSlide.BindPackageElement(OwnerSlide!, PackageTree!, PackageNode);
            _element = element;
            return element;
        }

        // The SDK part of a layout or master shape held in the package, opening the SDK when it is not open.
        private OpenXmlPart? PackagePartForSdk() {
            if (PackagePart.Current.SdkPart == null && PackagePart.Document is { } document) _ = document.OpenXmlDocument;
            return PackagePart.Current.SdkPart as OpenXmlPart;
        }

        internal PowerPointShape AttachTo(PowerPointSlide slide, OpenXmlPartContainer? ownerPart = null) {
            OwnerSlide = slide;
            RelationshipOwnerPart = ownerPart as OpenXmlPart ?? RelationshipOwnerPart ?? slide.SlidePart;
            return this;
        }

        private static bool IsDrawingElement(OpenXmlElement element) =>
            element is Shape or ConnectionShape or Picture or GraphicFrame or GroupShape;
    }
}
