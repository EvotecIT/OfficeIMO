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
        private bool _standalone;

        internal PowerPointHeldNode(OfficeXmlNode element, PowerPointShape owner) {
            _element = element;
            // An element of a copy kept after its part went belongs to no part and never binds.
            _standalone = PowerPointDetachedCopies.IsStandalone(element.Tree);
            if (!_standalone && element.Tree != null && owner.OwnerSlide is { } slide) {
                _packageSlide = slide;
                _packagePart = slide.PartRef;
                slide.OwnerDocument?.RegisterPackageBinding(this);
            }
        }

        /// <summary>The element as it is held now: once the SDK opened, a package node is its SDK element.</summary>
        internal OfficeXmlNode Element {
            get {
                // A removed slide's part name can be reused by a new slide, so its content never binds; an element removed
                // from its table has no place in the SDK and stays as it is.
                if (!_standalone && _element.Tree != null && _element.IsAttached && !_packagePart.IsNone && _packagePart.Document is { IsSdkLoaded: true } && _packageSlide is not { RemovedFromPackage: true }) {
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
            if (!_standalone && element.Tree != null && !_packagePart.IsNone) return PowerPointSlide.BindPackageNode(_packagePart, element.Tree, element.Node);
            throw new InvalidOperationException("The element is not held in an SDK part.");
        }

        void IPackageSdkBinding.KeepDetached(PowerPointDetachedCopies copies) {
            if (!_standalone && copies.ShouldCopy(_element, _packageSlide is { RemovedFromPackage: true })) {
                _element = copies.Copy(_element);
                _standalone = true;
            }
        }

        void IPackageSdkBinding.BindSdk() {
            try {
                _ = Element;
            } catch (InvalidOperationException) {
                // An element the opened slide doesn't have where expected reports it when it is used.
            }
        }
    }

    /// <summary>
    /// Copies of elements removed from their part (or whose slide was removed) taken before the package lets go of them.
    /// Wrappers of the same subtree share one copy, so they stay one piece of content, and nothing in a copy binds.
    /// </summary>
    internal sealed class PowerPointDetachedCopies {
        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<OfficeXmlTree, object> StandaloneTrees = new();
        private readonly Dictionary<(OfficeXmlTree Tree, int Root), OfficeXmlNode> _copies = new();

        /// <summary>Whether a tree is such a copy.</summary>
        internal static bool IsStandalone(OfficeXmlTree? tree) => tree != null && StandaloneTrees.TryGetValue(tree, out _);

        /// <summary>Whether a wrapper's element needs a copy: a package node that left its part, or whose slide was removed.</summary>
        internal bool ShouldCopy(OfficeXmlNode element, bool ownerRemoved) =>
            element.Tree != null && !IsStandalone(element.Tree) && (ownerRemoved || !element.IsAttached);

        /// <summary>The element's place in the shared copy of the subtree (or the whole part) it is in.</summary>
        internal OfficeXmlNode Copy(OfficeXmlNode element) {
            var path = new List<int>();
            OfficeXmlNode root = element;
            for (OfficeXmlNode parent = root.Parent; !parent.IsNone; parent = root.Parent) {
                path.Add(ElementIndex(parent, root));
                root = parent;
            }
            if (!_copies.TryGetValue((root.Tree!, root.Node), out OfficeXmlNode copy)) {
                byte[] markup = root.OuterXml();
                OfficeXmlTree tree = OfficeXmlTree.Parse(markup, markup.Length, root.Tree!.Namespaces);
                StandaloneTrees.Add(tree, tree);
                copy = new OfficeXmlNode(tree, tree.Root);
                _copies.Add((root.Tree!, root.Node), copy);
            }
            for (int level = path.Count - 1; level >= 0; level--) copy = NthElement(copy, path[level]);
            return copy;
        }

        private static int ElementIndex(OfficeXmlNode parent, OfficeXmlNode child) {
            int index = 0;
            foreach (OfficeXmlNode candidate in parent.Children()) {
                if (candidate.Equals(child)) return index;
                index++;
            }
            return -1;
        }

        private static OfficeXmlNode NthElement(OfficeXmlNode parent, int index) {
            foreach (OfficeXmlNode candidate in parent.Children()) {
                if (index-- == 0) return candidate;
            }
            return default;
        }
    }
}
