using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Shape members read and edited on the slide's package tree while the SDK is not open. Each returns false when it
    /// can't answer exactly as the SDK element would, and the member then binds the shape and reads it through the SDK.
    /// </summary>
    public abstract partial class PowerPointShape {
        /// <summary>
        /// The shape's element as it is held now: a node of its slide's package tree while the shape works there, the SDK
        /// element otherwise (binding it opens the SDK for a package shape the SDK does not hold yet).
        /// </summary>
        internal OfficeXmlNode XmlElement =>
            TryGetPackageElement(out OfficeXmlTree tree, out int node) ? new OfficeXmlNode(tree, node) : OfficeXmlNode.FromSdk(SdkElementObject);

        /// <summary>The part that holds the shape: its slide's part, or the layout, master or notes part it was read from.</summary>
        internal PowerPointPartRef OwnerPartRef {
            get {
                if (TryGetPackageElement(out _, out _)) return OwnerSlide!.PartRef;
                PowerPointPartRef current = PowerPointPartRef.FromSdk(OpenXmlSdk.OfficeSdkXml.PartOf(SdkElementObject));
                return !current.IsNone ? current : PowerPointPartRef.FromSdk(SdkRelationshipOwnerPart ?? OwnerSlide?.SdkSlidePart);
            }
        }

        /// <summary>The slide the shape is on, or none for a shape of a layout or master.</summary>
        internal PowerPointPartRef SlidePartRef {
            get {
                PowerPointPartRef owner = OwnerPartRef;
                return owner.Kind == PowerPointPartKind.Slide ? owner : OwnerSlide?.PartRef ?? default;
            }
        }

        /// <summary>The shape's package tree and element while it works without the SDK, or false once it is bound.</summary>
        internal bool TryGetPackageElement(out OfficeXmlTree tree, out int node) {
            tree = PackageTree!;
            node = PackageNode;
            // Once the SDK opened the package, the tree is gone and the shape binds to the SDK element instead.
            return IsPackageBacked && tree != null && OwnerSlide is { RemovedFromPackage: false } && OwnerSlide.OwnerDocument?.IsSdkLoaded == false;
        }

        /// <summary>
        /// The package to edit the shape in: false when the document can't be edited without the SDK, and the member then
        /// edits the shape through it.
        /// </summary>
        internal bool TryEditPackage(out OfficeXmlTree tree, out int node) {
            node = PackageNode;
            tree = PackageTree!;
            return TryGetPackageElement(out tree, out node) && OwnerSlide!.OwnerDocument!.CanEditPackage;
        }

        // The non-visual properties container of each drawing element, and the drawing-properties child the model adds
        // with it.
        private static string? NonVisualContainer(string localName) => localName switch {
            "sp" => "nvSpPr",
            "cxnSp" => "nvCxnSpPr",
            "pic" => "nvPicPr",
            "graphicFrame" => "nvGraphicFramePr",
            "grpSp" => "nvGrpSpPr",
            _ => null
        };

        private static string NonVisualDrawingChild(string localName) => localName switch {
            "sp" => "cNvSpPr",
            "cxnSp" => "cNvCxnSpPr",
            "pic" => "cNvPicPr",
            "graphicFrame" => "cNvGraphicFramePr",
            _ => "cNvGrpSpPr"
        };

        /// <summary>
        /// The shape's <c>p:cNvPr</c>; with <paramref name="create"/> the non-visual properties are added as the model adds
        /// them when they are missing. None for an element that has no non-visual properties.
        /// </summary>
        internal OfficeXmlNode NonVisualDrawingElement(bool create) {
            OfficeXmlNode element = XmlElement;
            if (element.NamespaceId != OfficeXmlNamespaces.PresentationML || NonVisualContainer(element.LocalName) is not { } containerName) return default;
            OfficeXmlNode container = element.Child(OfficeXmlNamespaces.PresentationML, containerName);
            if (container.IsNone) {
                if (!create) return default;
                OfficeXmlNode created = element.CreateElement(OfficeXmlNamespaces.PresentationML, containerName);
                created.Append(created.CreateElement(OfficeXmlNamespaces.PresentationML, "cNvPr"));
                OfficeXmlNode drawing = created.Append(created.CreateElement(OfficeXmlNamespaces.PresentationML, NonVisualDrawingChild(element.LocalName)));
                if (element.LocalName == "sp") drawing.Append(drawing.CreateElement(OfficeXmlNamespaces.DrawingML, "spLocks")).SetAttribute("noGrp", "1");
                created.Append(created.CreateElement(OfficeXmlNamespaces.PresentationML, "nvPr"));
                container = element.InsertBefore(created, element.FirstChild);
            }
            OfficeXmlNode properties = container.Child(OfficeXmlNamespaces.PresentationML, "cNvPr");
            if (properties.IsNone && create) properties = container.InsertBefore(container.CreateElement(OfficeXmlNamespaces.PresentationML, "cNvPr"), container.FirstChild);
            return properties;
        }

        /// <summary>The shape's <c>p:ph</c>, or none.</summary>
        internal OfficeXmlNode ShapePlaceholderElement {
            get {
                OfficeXmlNode element = XmlElement;
                if (element.NamespaceId != OfficeXmlNamespaces.PresentationML || element.LocalName == "grpSp" || NonVisualContainer(element.LocalName) is not { } containerName) return default;
                return element.Child(OfficeXmlNamespaces.PresentationML, containerName).Child(OfficeXmlNamespaces.PresentationML, "nvPr").Child(OfficeXmlNamespaces.PresentationML, "ph");
            }
        }

        // A picture that plays audio or video from a file.
        private static bool IsMediaPicture(OfficeXmlNode picture) {
            OfficeXmlNode application = picture.Child(OfficeXmlNamespaces.PresentationML, "nvPicPr").Child(OfficeXmlNamespaces.PresentationML, "nvPr");
            return ContainsMedia(application);
        }

        private static bool ContainsMedia(OfficeXmlNode element) {
            foreach (OfficeXmlNode child in element.Children()) {
                if (child.Is(OfficeXmlNamespaces.DrawingML, "audioFile") || child.Is(OfficeXmlNamespaces.DrawingML, "videoFile") || ContainsMedia(child)) return true;
            }
            return false;
        }

        /// <summary>Records that the shape's package tree changed, so the next save or SDK open writes it.</summary>
        internal void MarkPackageEdited() => OwnerSlide?.OwnerDocument?.MarkPackageEdited();
    }
}
