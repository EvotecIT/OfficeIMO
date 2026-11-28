using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// An element below a shape that a wrapper holds, such as a table row or cell: a node of the slide's package tree while
    /// the SDK is not open, which binds to the same element in the SDK when it opens, or an SDK element.
    /// </summary>
    internal sealed class PowerPointHeldNode : IPackageSdkBinding {
        private OfficeXmlNode _element;
        private readonly PowerPointPartRef _packagePart;
        private readonly PowerPointSlide? _packageSlide;
        // Whether the element is a standalone copy of one removed from its part.
        // Whether the element was removed before the package that held it went.
        private bool _removed;

        internal PowerPointHeldNode(OfficeXmlNode element, PowerPointShape owner) {
            _element = element;
            if (element.Tree == null) return;
            _packagePart = owner.OwnerPartRef;
            _packageSlide = owner.OwnerSlide;
            _packagePart.Document?.RegisterPackageBinding(this);
        }

        /// <summary>The element as it is held now: once the SDK opened, a package node is its SDK element.</summary>
        internal OfficeXmlNode Element {
            get {
                ThrowIfRemoved();
                // A removed slide's part name can be reused by a new slide, so its content never binds; an element removed
                // from its table has no place in the SDK and stays as it is.
                if (_element.Tree != null && _element.IsAttached && !_packagePart.IsNone && _packagePart.Document is { IsSdkLoaded: true } && _packageSlide is not { RemovedFromPackage: true }) {
                    _element = OfficeXmlNode.FromSdk(PowerPointSlide.BindPackageNode(_packagePart, _element.Tree, _element.Node));
                }
                return _element;
            }
        }

        /// <summary>
        /// The SDK element, for code that still works on SDK elements; a package node opens the SDK and binds to the same
        /// element there.
        /// </summary>
        internal object RequireSdkElement() {
            OfficeXmlNode element = Element;
            if (element.SdkElement is { } sdk) return sdk;
            if (_packageSlide is { RemovedFromPackage: true }) throw new InvalidOperationException("The slide is no longer part of the presentation.");
            if (element.Tree != null && !_packagePart.IsNone) return PowerPointSlide.BindPackageNode(_packagePart, element.Tree, element.Node);
            throw new InvalidOperationException("The element is not held in an SDK part.");
        }

        void IPackageSdkBinding.InvalidateRemoved(bool sdkOpening) {
            if (!_removed && _element.Tree != null && (OwnerSlideRemoved || sdkOpening && !_element.IsAttached)) _removed = true;
        }

        private void ThrowIfRemoved() {
            if (_removed) throw new InvalidOperationException("The content was removed from the presentation.");
        }

        private bool OwnerSlideRemoved => _packageSlide is { RemovedFromPackage: true };

        void IPackageSdkBinding.BindSdk() {
            try {
                _ = Element;
            } catch (InvalidOperationException) {
                // An element the opened slide doesn't have where expected reports it when it is used.
            }
        }
    }
}
