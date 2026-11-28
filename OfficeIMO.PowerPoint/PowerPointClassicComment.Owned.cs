using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointClassicComment {
        private const int PresentationNamespace = OfficeXmlNamespaces.PresentationML;

        private OfficeXmlNode AttachedOwnedComment() {
            if (_ownedComments == null || _presentation.IsSdkLoaded) return default;
            _presentation.ThrowIfDisposedForCommentApi();
            if (_slide.RemovedFromPackage) {
                throw new InvalidOperationException("The comment's slide is no longer attached to the presentation.");
            }
            OfficeXmlNode comment = _ownedComments.CommentAt(_ownedIndex);
            if (!comment.IsAttached) throw new InvalidOperationException("The classic comment is no longer attached to the presentation.");
            return comment;
        }

        private bool TryGetEditableComment(out OfficeXmlNode comment) {
            comment = AttachedOwnedComment();
            if (!comment.IsAttached || comment.Tree is not { } tree || _presentation.EditablePackage is not { } package ||
                _presentation.GetPackageXml(_ownedComments!.PartName) != tree || !_ownedComments.CanEditMarkup(package, tree)) return false;

            int previousRank = -1;
            for (int node = tree.FirstChildOf(comment.Node); node != OfficeXmlTree.NoNode; node = tree.NextSiblingOf(node)) {
                if (tree.KindOf(node) != OfficeXmlTreeNodeKind.Element) return false;
                var child = new OfficeXmlNode(tree, node);
                int rank = child.LocalName switch {
                    "pos" => 0,
                    "text" => 1,
                    "extLst" => 2,
                    _ => -1
                };
                if (rank <= previousRank || !child.Is(PresentationNamespace, child.LocalName) ||
                    tree.UriIdOf(child.Node) != OfficeXmlNamespaces.TransitionalUriId(PresentationNamespace)) return false;
                previousRank = rank;
            }
            return true;
        }

        private bool TryReassignOwnedComment(PowerPointCommentAuthor author) =>
            TryGetEditableComment(out OfficeXmlNode comment) &&
            _presentation.TryReassignPackageClassicComment(_slide, _ownedComments!.PartName, comment, author);

        private bool TryRemoveOwnedComment() =>
            TryGetEditableComment(out OfficeXmlNode comment) &&
            _presentation.TryRemovePackageClassicComment(_slide, _ownedComments!.PartName, comment);

        private bool TrySetOwnedText(string value) {
            if (value.IndexOf('\r') >= 0 || !TryGetEditableComment(out OfficeXmlNode comment)) return false;
            OfficeXmlNode previous = comment.Child(PresentationNamespace, "text");
            OfficeXmlNode replacement = comment.CreateElement(PresentationNamespace, "text");
            replacement.SetInnerText(value);
            comment.InsertBefore(replacement, previous.IsNone ? comment.Child(PresentationNamespace, "extLst") : previous);
            if (!previous.IsNone) previous.Remove();
            _presentation.MarkPackageEdited();
            return true;
        }

        private bool TrySetOwnedCreated(DateTime value) {
            if (value.Kind != DateTimeKind.Utc || !TryGetEditableComment(out OfficeXmlNode comment)) return false;
            comment.SetAttribute("dt", XmlConvert.ToString(value, XmlDateTimeSerializationMode.Utc));
            _presentation.MarkPackageEdited();
            return true;
        }

        private bool TrySetOwnedPosition(string axis, long value) {
            if (!TryGetEditableComment(out OfficeXmlNode comment)) return false;
            OfficeXmlNode position = comment.Child(PresentationNamespace, "pos");
            if (position.IsNone) {
                position = comment.InsertBefore(comment.CreateElement(PresentationNamespace, "pos"), comment.FirstChild);
                position.SetInt64Attribute("x", 0L);
                position.SetInt64Attribute("y", 0L);
            }
            position.SetInt64Attribute(axis, value);
            _presentation.MarkPackageEdited();
            return true;
        }

        private bool TryReadOwnedUInt(string attribute, out uint? value) {
            OfficeXmlNode comment = AttachedOwnedComment();
            string? text = comment.Attribute(attribute);
            value = text is { Length: > 0 } ? XmlConvert.ToUInt32(text) : null;
            return !comment.IsNone && text != string.Empty;
        }

        private string? ReadOwnedText() {
            OfficeXmlNode comment = AttachedOwnedComment();
            if (comment.IsNone) return null;
            OfficeXmlNode text = comment.Child(PresentationNamespace, "text");
            if (text.IsNone) return string.Empty;
            OfficeXmlTree tree = text.Tree!;
            int first = tree.FirstChildOf(text.Node);
            if (first == OfficeXmlTree.NoNode) return string.Empty;
            // The SDK reads split text at its first text segment. Let it handle segments and unusual leaf markup.
            if (tree.KindOf(first) != OfficeXmlTreeNodeKind.Text || tree.NextSiblingOf(first) != OfficeXmlTree.NoNode) return null;
            string value = tree.GetText(first);
            return string.IsNullOrWhiteSpace(value) ? null : value;
        }

        private bool TryReadOwnedCreated(out DateTime? created) {
            OfficeXmlNode comment = AttachedOwnedComment();
            string? text = comment.Attribute("dt");
            created = text is { Length: > 0 } ? XmlConvert.ToDateTime(text, XmlDateTimeSerializationMode.Utc) : null;
            return !comment.IsNone && text != string.Empty;
        }

        private long? ReadOwnedPosition(string axis) {
            OfficeXmlNode position = AttachedOwnedComment().Child(PresentationNamespace, "pos");
            if (position.IsNone) return null;
            string? text = position.Attribute(axis);
            if (text == string.Empty) return null;
            return text == null ? 0L : XmlConvert.ToInt64(text);
        }
    }

    /// <summary>Binds the comments held from one inventory to their corresponding SDK elements in one pass.</summary>
    internal sealed class PowerPointClassicCommentListBinding : IPackageSdkBinding {
        private readonly PowerPointSlide _slide;
        private readonly string _partName;
        private IReadOnlyList<OfficeXmlNode>? _comments;
        private OfficeXmlNode _originalRoot;
        private readonly int _originalVersion;
        private OfficeXmlNode[]? _sdkComments;
        private bool _removed;
        private bool? _editableMarkup;

        internal PowerPointClassicCommentListBinding(PowerPointDocument document, PowerPointSlide slide, string partName,
            IReadOnlyList<OfficeXmlNode> comments) {
            _slide = slide;
            _partName = partName;
            _comments = comments;
            _originalRoot = comments[0].Parent;
            _originalVersion = _originalRoot.Tree!.Version;
            document.RegisterPackageBinding(this);
        }

        internal string PartName => _partName;

        internal bool CanEditMarkup(OfficeOpcEditablePackage package, OfficeXmlTree tree) =>
            _editableMarkup ??= !tree.HasCommentsOrInstructions() && HasValidXml(package);

        private bool HasValidXml(OfficeOpcEditablePackage package) {
            try {
                // Unchanged extension bytes are copied verbatim by the tree writer. Decode them before the first edit.
                using Stream source = package.OpenPartContentStream(_partName, default);
                PowerPointXmlReader.ValidatePackagePartXml(source);
                return true;
            } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
                return false;
            }
        }

        internal OfficeXmlNode CommentAt(int index) => _comments?[index] ?? default;

        internal OfficeXmlNode SdkCommentAt(int index) => _sdkComments?[index] ?? default;

        void IPackageSdkBinding.InvalidateRemoved(bool sdkOpening) => _removed |= _slide.RemovedFromPackage;

        void IPackageSdkBinding.BindSdk() {
            try {
                if (_removed) return;
                if (_comments is not { } original || _originalRoot.Tree is not { } tree) return;
                OfficeXmlNode[] comments = PowerPointClassicComment.ReadSdkCommentsForBinding(_slide, _partName);
                if (comments.Length == original.Count && tree.Version == _originalVersion) {
                    _sdkComments = comments;
                    return;
                }

                // An owned edit may have added or removed comments since the caller received this inventory.
                var originalIndexes = new Dictionary<OfficeXmlNode, int>(original.Count);
                for (int index = 0; index < original.Count; index++) originalIndexes.Add(original[index], index);
                var mapped = new OfficeXmlNode[original.Count];
                int sdkIndex = 0;
                foreach (OfficeXmlNode current in _originalRoot.Children(OfficeXmlNamespaces.PresentationML, "cm")) {
                    if (sdkIndex >= comments.Length) return;
                    if (originalIndexes.TryGetValue(current, out int index)) mapped[index] = comments[sdkIndex];
                    sdkIndex++;
                }
                if (sdkIndex == comments.Length) _sdkComments = mapped;
            } catch (InvalidOperationException) {
                // A wrapper reports when its original comment did not survive the handoff.
            } finally {
                _comments = null;
                _originalRoot = default;
            }
        }
    }
}
